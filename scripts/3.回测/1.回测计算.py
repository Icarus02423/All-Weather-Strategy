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


def _metrics(returns, annualization, risk_free_rate, nav_values, initial_nav, dates, baseline_date):
    returns = np.asarray(returns, dtype=float)
    count = len(returns)
    nav_values = np.asarray(nav_values, dtype=float)
    total = float(nav_values[-1] / initial_nav - 1)
    # 先减 1 会把很小但正的财富比舍入成 -1，不能再对 total 做 log1p。
    try:
        annual_return = math.expm1((math.log(float(nav_values[-1])) - math.log(initial_nav)) * annualization / count)
    except OverflowError:
        annual_return = None
    vol = float(np.std(returns, ddof=1) * math.sqrt(annualization)) if count > 1 else None
    risk_free_daily = math.expm1(math.log1p(risk_free_rate) / annualization)
    excess_returns = returns - risk_free_daily
    annual_excess = float(np.mean(excess_returns) * annualization)
    sharpe = annual_excess / vol if vol is not None and vol > 1e-14 else None
    # ETF 口径：所有观察日均进入均方分母，上行日的下行偏差为零。
    downside = float(np.sqrt(np.mean(np.minimum(excess_returns, 0.) ** 2)) * math.sqrt(annualization))
    sortino = annual_excess / downside if downside > 0 else None
    if sortino is not None and not math.isfinite(sortino):
        sortino = None
    nav_path = np.r_[initial_nav, nav_values]
    peak = np.maximum.accumulate(nav_path)
    drawdown_path = nav_path / peak - 1
    trough_position = int(np.argmin(drawdown_path))
    maximum_drawdown = float(-drawdown_path[trough_position])
    drawdown_start = drawdown_end = None
    if maximum_drawdown > 0:
        # 与 ETF 一致：重复峰值取谷值之前最后一次；重复最低回撤取首次谷值。
        peak_position = int(np.flatnonzero(nav_path[:trough_position + 1] == peak[trough_position])[-1])
        date_path = [pd.Timestamp(baseline_date), *[pd.Timestamp(day) for day in dates]]
        drawdown_start, drawdown_end = date_path[peak_position], date_path[trough_position]
    return {"observations": count, "total_return": total, "annualized_return": annual_return,
            "annualized_volatility": vol, "sharpe_ratio": sharpe, "sortino_ratio": sortino,
            "max_drawdown": maximum_drawdown, "max_drawdown_start": drawdown_start,
            "max_drawdown_end": drawdown_end,
            "calmar_ratio": annual_return / maximum_drawdown if annual_return is not None and maximum_drawdown > 1e-14 else None,
            "annualized_return_status": "defined" if annual_return is not None else "outside_float_range",
            "positive_day_fraction": float(np.mean(returns > 0))}


def _turnover_metrics(rows, annualization):
    # 仅排除全回测的首次实际建仓；保留其余现金日和非调仓日的零换手。
    retained = rows.loc[~rows["is_initial_entry"], "one_way_turnover"]
    amount = float(retained.sum())
    return {"mean_one_way_turnover": amount / len(retained) if len(retained) else None,
            "annualized_one_way_turnover": amount / (len(rows) / annualization)}


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


def _benchmark_path(benchmark_prices, calendar, baseline):
    if not isinstance(benchmark_prices, pd.Series) or benchmark_prices.empty:
        raise BacktestError("Benchmark prices must be a nonempty Series")
    index = benchmark_prices.index
    if (not isinstance(index, pd.DatetimeIndex) or index.hasnans or index.tz is not None
            or not index.is_unique or not index.is_monotonic_increasing or not index.equals(index.normalize())):
        raise BacktestError("Benchmark dates must be unique, ascending daily dates")
    dates = pd.DatetimeIndex([baseline, *calendar])
    if not dates.isin(index).all():
        raise BacktestError("Benchmark is missing the common baseline or backtest dates; no filling or truncation")
    values = np.array([_number(v, "benchmark close") for v in benchmark_prices.loc[dates]])
    if np.any(values <= 0):
        raise BacktestError("Benchmark closes must be positive")
    return values / values[0]


def _relative_metrics(rows, annualization, risk_free_rate, baseline):
    strategy_initial = float(rows["nav_start"].iloc[0])
    benchmark_initial = float(rows["benchmark_nav_start"].iloc[0])
    benchmark = _metrics(rows["benchmark_return"], annualization, risk_free_rate, rows["benchmark_nav"],
                         benchmark_initial, rows["date"], baseline)
    strategy = _metrics(rows["net_return"], annualization, risk_free_rate, rows["nav"],
                        strategy_initial, rows["date"], baseline)
    relative_nav = (rows["nav"] / strategy_initial) / (rows["benchmark_nav"] / benchmark_initial)
    relative_returns = (1 + rows["net_return"]) / (1 + rows["benchmark_return"]) - 1
    relative = _metrics(relative_returns, annualization, 0., relative_nav, 1., rows["date"], baseline)
    active = rows["net_return"] - rows["benchmark_return"]
    tracking = float(active.std(ddof=1) * math.sqrt(annualization)) if len(rows) > 1 else None
    annual_difference = (strategy["annualized_return"] - benchmark["annualized_return"]
                         if strategy["annualized_return"] is not None and benchmark["annualized_return"] is not None else None)
    return {"benchmark_return": benchmark["total_return"],
            "benchmark_annualized_return": benchmark["annualized_return"],
            "benchmark_annualized_volatility": benchmark["annualized_volatility"],
            "benchmark_max_drawdown": benchmark["max_drawdown"],
            "benchmark_max_drawdown_start": benchmark["max_drawdown_start"],
            "benchmark_max_drawdown_end": benchmark["max_drawdown_end"],
            "excess_return": strategy["total_return"] - benchmark["total_return"],
            "annualized_excess_return": annual_difference, "relative_total_return": relative["total_return"],
            "tracking_error": tracking,
            "information_ratio": float(active.mean() * annualization / tracking) if tracking is not None and tracking > 1e-14 else None,
            "excess_max_drawdown": relative["max_drawdown"],
            "excess_max_drawdown_start": relative["max_drawdown_start"],
            "excess_max_drawdown_end": relative["max_drawdown_end"]}


def _period_returns(daily, frequency, annualization, risk_free_rate, baseline_date, baseline_policy):
    records = []
    labels = daily["date"].dt.to_period(frequency)
    for period, rows in daily.groupby(labels, sort=True):
        initial_nav = float(rows["nav_start"].iloc[0])
        record = {"period": str(period), "first_date": rows["date"].iloc[0], "last_date": rows["date"].iloc[-1],
                  "observations": len(rows), "return": float(rows["nav"].iloc[-1] / initial_nav - 1),
                  "trading_pnl": float(rows["trading_pnl"].sum()), "total_cost": float(rows["total_cost"].sum())}
        if frequency == "Y":
            first_position = daily.index.get_loc(rows.index[0])
            year_baseline = daily["date"].iloc[first_position - 1] if first_position else baseline_date
            year_policy = "previous_backtest_observation" if first_position else baseline_policy
            stats = _metrics(rows["net_return"], annualization, risk_free_rate, rows["nav"],
                             initial_nav, rows["date"], year_baseline)
            record.update({key: stats[key] for key in (
                "annualized_volatility", "sharpe_ratio", "sortino_ratio", "max_drawdown",
                "max_drawdown_start", "max_drawdown_end", "calmar_ratio")})
            record.update({"annualized_one_way_turnover": _turnover_metrics(rows, annualization)["annualized_one_way_turnover"],
                           "baseline_date": year_baseline, "baseline_date_policy": year_policy})
            record.update(_relative_metrics(rows, annualization, risk_free_rate, year_baseline))
        records.append(record)
    return pd.DataFrame(records)


