from __future__ import annotations

from collections.abc import Mapping
import math

import numpy as np
import pandas as pd

GROUPS = ("equity", "bonds", "gold", "other_commodities")
SUCCESS = {"target_achieved", "cap_limited"}
ENGINE_VERSION = "fixed_units_bond_margin_benchmark_1.0"


class BacktestError(ValueError):
    """输入或记账失败；禁止将该次运行当作有效回测。"""


def _number(value, name, minimum=None):
    if isinstance(value, (bool, str)) or not isinstance(value, (int, float, np.number)):
        raise BacktestError(f"{name} must be numeric")
    value = float(value)
    if not math.isfinite(value) or (minimum is not None and value < minimum):
        raise BacktestError(f"Invalid {name}: {value}")
    return value


def _date(value, name):
    if value is None or isinstance(value, (bool, int, float, np.number)):
        raise BacktestError(f"Invalid {name}")
    try:
        result = pd.Timestamp(value)
    except (ValueError, TypeError) as exc:
        raise BacktestError(f"Invalid {name}") from exc
    if pd.isna(result) or result.tzinfo is not None or result != result.normalize():
        raise BacktestError(f"Invalid {name}")
    return result


def _validate_prices(prices):
    if not isinstance(prices, pd.DataFrame) or prices.empty:
        raise BacktestError("Prices must be a nonempty DataFrame")
    index = prices.index
    if (not isinstance(index, pd.DatetimeIndex) or index.hasnans or index.tz is not None
            or not index.equals(index.normalize()) or not index.is_unique
            or not index.is_monotonic_increasing):
        raise BacktestError("Price dates must be unique, increasing daily dates")
    if not prices.columns.is_unique or any(not isinstance(c, str) or not c.strip() or c != c.strip() for c in prices.columns):
        raise BacktestError("Price columns must be unique instrument codes")
    try:
        return prices.astype(float)
    except (TypeError, ValueError) as exc:
        raise BacktestError("Prices must be numeric; unused missing prices may remain NaN") from exc


def _prepare_signals(signals, prices):
    if not signals:
        raise BacktestError("No optimization signals supplied")
    prepared, numbers, dates, executions, labels = [], set(), set(), set(), {}
    for source in signals:
        if not isinstance(source, Mapping) or source.get("status") not in SUCCESS:
            raise BacktestError("Only successful, validated optimization signals may be backtested")
        number = source.get("rebalance_no")
        if isinstance(number, bool) or not isinstance(number, (int, np.integer)) or number < 1 or number in numbers:
            raise BacktestError("Rebalance numbers must be distinct positive integers")
        day = _date(source.get("signal_date"), "signal_date")
        if day not in prices.index or day in dates:
            raise BacktestError("Signal dates must be distinct observed price dates")
        execution = source.get("execution_date")
        location = prices.index.get_loc(day)
        if execution is None:
            if location != len(prices.index) - 1:
                raise BacktestError("A signal without execution date must be the final observed date")
        else:
            execution = _date(execution, "execution_date")
            if location + 1 >= len(prices.index) or execution != prices.index[location + 1] or execution in executions:
                raise BacktestError("Execution must be the unique next observed date after its signal")
            executions.add(execution)
        weights, groups = source.get("exposures"), source.get("groups")
        if not isinstance(weights, Mapping) or not weights or not isinstance(groups, Mapping) or set(weights) != set(groups):
            raise BacktestError("Each signal needs exposure and group mappings for the same instruments")
        checked = {}
        for code, value in weights.items():
            if code not in prices.columns or groups[code] not in GROUPS:
                raise BacktestError(f"Unknown instrument/group: {code}")
            weight = _number(value, f"exposure {code}")
            if weight <= 0:
                raise BacktestError("Included instrument exposures must be strictly positive")
            if code in labels and labels[code] != groups[code]:
                raise BacktestError(f"Instrument changes asset class: {code}")
            labels[code] = groups[code]
            checked[code] = weight
        target = _number(source.get("target_volatility"), "target_volatility")
        forecast = _number(source.get("actual_volatility"), "actual_volatility")
        if target <= 0 or forecast <= 0:
            raise BacktestError("Target and forecast volatility must be positive")
        floor = _number(source.get("bond_margin_floor"), "bond_margin_floor")
        limit = _number(source.get("bond_leverage_cap"), "bond_leverage_cap")
        margin = _number(source.get("bond_margin_rate"), "bond_margin_rate")
        capital = source.get("capital_weights")
        if floor <= 0 or limit <= 0 or margin <= 0 or not math.isclose(floor * limit, 1., abs_tol=1e-10, rel_tol=1e-10):
            raise BacktestError("Invalid bond margin/leverage parameters")
        if set(groups.values()) != set(GROUPS) or not isinstance(capital, Mapping) or set(capital) != set(checked):
            raise BacktestError("Capital weights must cover exactly the scheduled instruments and four groups")
        checked_capital = {c: _number(capital[c], f"capital_weight {c}", 0) for c in checked}
        for c, exposure in checked.items():
            expected_capital = exposure * (margin if groups[c] == "bonds" else 1.)
            if not math.isclose(checked_capital[c], expected_capital, rel_tol=1e-9, abs_tol=1e-10):
                raise BacktestError(f"Exposure and capital weight disagree: {c}")
        bond = math.fsum(v for c, v in checked.items() if groups[c] == "bonds")
        non_bond = math.fsum(v for c, v in checked.items() if groups[c] != "bonds")
        if (abs(math.fsum(checked_capital.values()) - 1) > 1e-9 or non_bond + floor * bond > 1 + 1e-9
                or (floor - margin) * bond > 1e-9):
            raise BacktestError("Target allocation violates the capital budget or margin floor")
        prepared.append({**source, "rebalance_no": int(number), "signal_date": day,
                         "execution_date": execution, "exposures": checked, "groups": dict(groups),
                         "capital_weights": checked_capital, "bond_margin_floor": floor,
                         "bond_leverage_cap": limit, "bond_margin_rate": margin,
                         "target_volatility": target, "actual_volatility": forecast})
        numbers.add(number)
        dates.add(day)
    prepared.sort(key=lambda s: s["signal_date"])
    if len({s["target_volatility"] for s in prepared}) != 1:
        raise BacktestError("One backtest must use one target volatility")
    if len({(s["bond_margin_floor"], s["bond_leverage_cap"]) for s in prepared}) != 1:
        raise BacktestError("One backtest must use one bond leverage cap")
    if any(b["rebalance_no"] <= a["rebalance_no"] for a, b in zip(prepared, prepared[1:])):
        raise BacktestError("Rebalance numbers must increase with signal dates")
    return prepared, labels


def _post_cost_nav(nav_before, old_notional, targets, rate):
    if rate == 0:
        return nav_before
    if rate * sum(targets.values()) >= 1:
        raise BacktestError("Trading cost times target gross exposure must be less than one")
    liquidation_cost = rate * sum(abs(v) for v in old_notional.values())
    if liquidation_cost >= nav_before:
        raise BacktestError("Insufficient equity to cover transaction costs")
    codes = set(old_notional) | set(targets)
    low, high = 0.0, nav_before
    for _ in range(100):
        middle = (low + high) / 2
        value = middle + rate * sum(abs(targets.get(c, 0.) * middle - old_notional.get(c, 0.)) for c in codes) - nav_before
        if value > 0:
            high = middle
        else:
            low = middle
    value = (low + high) / 2
    if value <= 0 or not math.isfinite(value):
        raise BacktestError("Nonpositive equity after trading costs")
    return value






def _funding_state(nav, notionals, groups, floor):
    non_bond = math.fsum(v for c, v in notionals.items() if groups[c] != "bonds")
    bond = math.fsum(v for c, v in notionals.items() if groups[c] == "bonds")
    invested = bond > 0
    capital = nav - non_bond if invested else None
    required = floor * bond if invested else 0.
    buffer = capital - required if invested else None
    margin = capital / bond if invested else None
    leverage = bond / capital if invested and capital > 0 else None
    breach = bool(invested and buffer < -1e-9 * nav)
    depleted = bool(invested and capital <= 0)
    return {"bond_notional": bond, "bond_exposure": bond / nav,
            "bond_capital": capital, "bond_capital_weight": capital / nav if invested else None,
            "bond_margin_rate": margin, "bond_leverage": leverage,
            "minimum_bond_capital": required, "capital_buffer": buffer,
            "capital_shortfall": max(0., -buffer) if invested else 0.,
            "capital_used_at_floor": (non_bond + required) / nav,
            "capital_usage": (non_bond + margin * bond) / nav if invested else None,
            "margin_breach": breach, "bond_capital_depleted": depleted,
            "funding_state": "not_invested" if not invested else
                "bond_capital_depleted" if depleted else "below_floor" if breach else "within_limit"}








