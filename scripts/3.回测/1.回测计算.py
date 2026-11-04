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


def backtest(prices, signals, *, benchmark_prices, start_date=None, end_date=None, initial_nav=1.,
             fee_bps=10., slippage_bps=0., annualization=252, risk_free_rate=0., rolling_window=63):
    prices = _validate_prices(prices)
    signals, labels = _prepare_signals(list(signals), prices)
    initial_nav = _number(initial_nav, "initial_nav")
    fee_bps = _number(fee_bps, "fee_bps", 0)
    slippage_bps = _number(slippage_bps, "slippage_bps", 0)
    risk_free_rate = _number(risk_free_rate, "risk_free_rate")
    if initial_nav <= 0 or risk_free_rate <= -1:
        raise BacktestError("Initial NAV must be positive and annual risk-free rate greater than -1")
    for value, name in ((annualization, "annualization"), (rolling_window, "rolling_window")):
        if isinstance(value, bool) or not isinstance(value, (int, np.integer)) or value < 2:
            raise BacktestError(f"{name} must be an integer of at least two")
    available = [s["execution_date"] for s in signals if s["execution_date"] is not None]
    if not available:
        raise BacktestError("No executable optimization signal")
    start = _date(start_date, "start_date") if start_date is not None else min(available)
    end = _date(end_date, "end_date") if end_date is not None else prices.index[-1]
    if start > end or start < prices.index[0] or end > prices.index[-1]:
        raise BacktestError("Requested dates must be ordered and covered by the price calendar")
    calendar = prices.index[(prices.index >= start) & (prices.index <= end)]
    if len(calendar) == 0:
        raise BacktestError("No observed dates in the requested interval")
    first_position = prices.index.get_loc(calendar[0])
    if first_position == 0:
        raise BacktestError("A previous observed date is required for the common benchmark baseline")
    baseline_date = prices.index[first_position - 1]
    benchmark_nav = _benchmark_path(benchmark_prices, calendar, baseline_date)
    margin_floor, leverage_cap = signals[0]["bond_margin_floor"], signals[0]["bond_leverage_cap"]
    events = {s["execution_date"]: s for s in signals if s["execution_date"] is not None and start <= s["execution_date"] <= end}
    if not events:
        raise BacktestError("No execution in the requested interval; refusing an all-cash strategy result")
    fee_rate, slip_rate = fee_bps / 10000., slippage_bps / 10000.
    rate = fee_rate + slip_rate
    nav, peak, units = initial_nav, initial_nav, {}
    daily_rows, instrument_rows, group_rows, trades = [], [], [], []
    allocation_rows = []
    capital_target_error = 0.
    accounting_error, contribution_error, target_error = 0., 0., 0.
    active_signal = None

    def price(day, code):
        value = float(prices.at[day, code])
        if not math.isfinite(value) or value <= 0:
            raise BacktestError(f"Missing/nonpositive/nonfinite required price: {day:%Y-%m-%d} {code}")
        return value

    for day in calendar:
        previous_nav = nav
        previous_units = units.copy()
        location = prices.index.get_loc(day)
        previous_day = prices.index[location - 1] if location > 0 else None
        pnl, previous_prices, current_prices = {}, {}, {}
        for code, quantity in previous_units.items():
            if previous_day is None:
                raise BacktestError("Held position has no previous observed price")
            previous_prices[code] = price(previous_day, code)
            current_prices[code] = price(day, code)
            pnl[code] = quantity * (current_prices[code] - previous_prices[code])
        trading_pnl = math.fsum(pnl.values())
        nav_before = previous_nav + trading_pnl
        if not math.isfinite(nav_before) or nav_before <= 0:
            raise BacktestError(f"Portfolio equity became nonpositive/nonfinite on {day:%Y-%m-%d}")
        funding_before = _funding_state(nav_before, {c: q * current_prices[c] for c, q in previous_units.items()}, labels, margin_floor)
        event = events.get(day)
        notionals, fees, slips = {}, {}, {}
        turnover = 0.
        if event is not None:
            codes = set(previous_units) | set(event["exposures"])
            for code in codes:
                current_prices[code] = price(day, code)
            old_notional = {c: previous_units.get(c, 0.) * current_prices[c] for c in codes}
            nav = _post_cost_nav(nav_before, old_notional, event["exposures"], rate)
            units = {c: weight * nav / current_prices[c] for c, weight in event["exposures"].items()}
            notionals = {c: (units.get(c, 0.) - previous_units.get(c, 0.)) * current_prices[c] for c in codes}
            fees = {c: abs(value) * fee_rate for c, value in notionals.items()}
            slips = {c: abs(value) * slip_rate for c, value in notionals.items()}
            turnover = math.fsum(abs(v) for v in notionals.values()) / nav_before
            target_error = max(target_error, max(abs(units[c] * current_prices[c] / nav - w) for c, w in event["exposures"].items()))
            funding_after = _funding_state(nav, {c: q * current_prices[c] for c, q in units.items()}, labels, margin_floor)
            capital_target_error = max(capital_target_error, abs(funding_after["capital_usage"] - 1),
                                       abs(funding_after["bond_margin_rate"] - event["bond_margin_rate"]) * funding_after["bond_exposure"])
            if funding_after["margin_breach"] or funding_after["bond_capital_depleted"]:
                raise BacktestError("Post-trade allocation violates the capital constraint")
            for code in event["exposures"]:
                allocation_rows.append({"date": day, "rebalance_no": event["rebalance_no"], "instrument_code": code,
                                        "asset_class": labels[code], "target_exposure": event["exposures"][code],
                                        "capital_weight": event["capital_weights"][code],
                                        "target_notional": event["exposures"][code] * nav,
                                        "capital_amount": event["capital_weights"][code] * nav,
                                        "bond_margin_rate": event["bond_margin_rate"]})
            active_signal = event
            trades.append({"date": day, "rebalance_no": event["rebalance_no"], "signal_date": event["signal_date"],
                           "optimization_status": event["status"], "target_volatility": event["target_volatility"],
                           "forecast_volatility": event["actual_volatility"], "nav_before_cost": nav_before, "nav_after_cost": nav,
                           "traded_notional": math.fsum(abs(v) for v in notionals.values()), "turnover": turnover,
                           "fee": math.fsum(fees.values()), "slippage": math.fsum(slips.values()),
                           "total_cost": math.fsum(fees.values()) + math.fsum(slips.values()),
                           "target_gross_exposure": math.fsum(event["exposures"].values()),
                           "target_non_bond_exposure": math.fsum(v for c, v in event["exposures"].items() if labels[c] != "bonds"),
                           "bond_margin_floor": margin_floor, "bond_leverage_cap": leverage_cap,
                           "target_bond_margin_rate": event["bond_margin_rate"],
                           **{f"{key}_before_trade": value for key, value in funding_before.items()}, **funding_after})
        else:
            nav = nav_before
        total_fee, total_slippage = math.fsum(fees.values()), math.fsum(slips.values())
        total_cost = total_fee + total_slippage
        accounting_error = max(accounting_error, abs(nav - (previous_nav + trading_pnl - total_cost)))
        net_return = nav / previous_nav - 1
        peak = max(peak, nav)
        exposure = {c: q * current_prices[c] / nav for c, q in units.items()}
        gross_exposure = math.fsum(exposure.values())
        non_bond = math.fsum(v for c, v in exposure.items() if labels[c] != "bonds")
        funding = _funding_state(nav, {c: q * current_prices[c] for c, q in units.items()}, labels, margin_floor)
        if not all(math.isfinite(v) for v in (net_return, nav, gross_exposure, non_bond)):
            raise BacktestError("Nonfinite portfolio accounting result")
        day_instruments = []
        for code in sorted(set(previous_units) | set(units)):
            quantity_before, quantity_after = previous_units.get(code, 0.), units.get(code, 0.)
            amount = pnl.get(code, 0.)
            fee, slippage = fees.get(code, 0.), slips.get(code, 0.)
            record = {"date": day, "instrument_code": code, "asset_class": labels[code],
                      "units_start": quantity_before, "units_end": quantity_after, "price": current_prices[code],
                      "previous_price": previous_prices.get(code),
                      "exposure_start": quantity_before * previous_prices[code] / previous_nav if quantity_before else 0.,
                      "exposure_end": exposure.get(code, 0.), "pnl": amount,
                      "return_contribution": amount / previous_nav,
                      "net_return_contribution": (amount - fee - slippage) / previous_nav,
                      "traded_notional": notionals.get(code, 0.), "fee": fee, "slippage": slippage,
                      "total_cost": fee + slippage}
            day_instruments.append(record)
            instrument_rows.append(record)
        contribution_error = max(contribution_error, abs(math.fsum(r["net_return_contribution"] for r in day_instruments) - net_return))
        for group in GROUPS:
            rows = [r for r in day_instruments if r["asset_class"] == group]
            group_rows.append({"date": day, "asset_class": group, **{column: math.fsum(r[column] for r in rows)
                              for column in ("exposure_start", "exposure_end", "pnl", "return_contribution", "net_return_contribution", "fee", "slippage", "total_cost")}})
        daily_rows.append({"date": day, "nav_start": previous_nav, "nav": nav, "normalized_nav": nav / initial_nav, "nav_before_cost": nav_before,
                           "gross_return": trading_pnl / previous_nav, "net_return": net_return,
                           "trading_pnl": trading_pnl, "fee": total_fee, "slippage": total_slippage, "total_cost": total_cost,
                           "drawdown": nav / peak - 1, "gross_exposure": gross_exposure, "non_bond_exposure": non_bond,
                           "turnover": turnover, **funding,
                           **{f"{key}_before_trade": value for key, value in funding_before.items()},
                           "bond_margin_floor": margin_floor, "bond_leverage_cap": leverage_cap,
                           "target_bond_margin_rate": active_signal["bond_margin_rate"] if active_signal else None,
                           "margin_breach_observed": funding_before["margin_breach"] or funding["margin_breach"],
                           "rebalance_no": event["rebalance_no"] if event is not None else None,
                           "held_signal_date": active_signal["signal_date"] if active_signal else None,
                           "target_volatility": signals[0]["target_volatility"],
                           "forecast_volatility_at_rebalance": active_signal["actual_volatility"] if active_signal else None})
    daily = pd.DataFrame(daily_rows)
    daily["benchmark_nav_start"], daily["benchmark_nav"] = benchmark_nav[:-1], benchmark_nav[1:]
    daily["benchmark_return"] = benchmark_nav[1:] / benchmark_nav[:-1] - 1
    daily["active_return"] = daily["net_return"] - daily["benchmark_return"]
    daily["relative_nav"] = daily["normalized_nav"] / daily["benchmark_nav"]
    daily["benchmark_drawdown"] = benchmark_nav[1:] / np.maximum.accumulate(benchmark_nav)[1:] - 1
    relative_path = np.r_[1., daily["relative_nav"]]
    daily["relative_drawdown"] = relative_path[1:] / np.maximum.accumulate(relative_path)[1:] - 1
    # 以下均为报表统计，不改动已计算的单位、交易、费用或净值。
    daily["one_way_turnover"] = daily["turnover"] / 2
    daily["cost_rate"] = daily["total_cost"] / daily["nav_before_cost"]
    daily["cumulative_cost_rate"] = 1 - (1 - daily["cost_rate"]).cumprod()
    daily["is_initial_entry"] = False
    first_entry = daily.index[daily["rebalance_no"].notna()][0]
    daily.loc[first_entry, "is_initial_entry"] = True
    instruments = pd.DataFrame(instrument_rows)
    counts = instruments.loc[instruments["units_end"] > 0].groupby("date").size()
    daily["holding_count"] = daily["date"].map(counts).fillna(0).astype(int)
    first_position = prices.index.get_loc(daily["date"].iloc[0])
    if first_position:
        baseline_date = prices.index[first_position - 1]
        baseline_policy = "previous_observed_price_date_initial_account_state"
    else:
        baseline_date = daily["date"].iloc[0] - pd.Timedelta(days=1)
        baseline_policy = "synthetic_previous_calendar_day_initial_account_marker_not_market_data"
    daily["realized_volatility_rolling"] = daily["net_return"].rolling(rolling_window, min_periods=rolling_window).std(ddof=1) * math.sqrt(annualization)
    signal_rows = []
    for signal in signals:
        execution = signal["execution_date"]
        if execution is None:
            action = "no_execution_date"
        elif execution < start:
            action = "before_start_not_inherited"
        elif execution > end:
            action = "after_end_not_executed"
        else:
            action = "executed"
        signal_rows.append({"rebalance_no": signal["rebalance_no"], "signal_date": signal["signal_date"],
                            "execution_date": execution, "optimization_status": signal["status"], "backtest_action": action,
                            "target_volatility": signal["target_volatility"], "forecast_volatility": signal["actual_volatility"],
                            "bond_leverage_cap": leverage_cap, "bond_margin_floor": margin_floor,
                            "target_bond_margin_rate": signal["bond_margin_rate"],
                            "source_path": signal.get("source_path"), "source_sha256": signal.get("source_sha256")})
    attribution_rows = []
    for level, key, frame in (("group", "asset_class", pd.DataFrame(group_rows)),
                              ("instrument", "instrument_code", pd.DataFrame(instrument_rows))):
        for name, rows in frame.groupby(key, sort=False):
            values = {column: float(rows[column].sum()) for column in ("pnl", "fee", "slippage", "total_cost")}
            net_pnl = values["pnl"] - values["total_cost"]
            attribution_rows.append({"level": level, "name": name, "trading_pnl": values["pnl"],
                                     "fee": values["fee"], "slippage": values["slippage"], "total_cost": values["total_cost"],
                                     "net_pnl": net_pnl, "contribution_to_total_return": net_pnl / initial_nav})
    metrics = _metrics(daily["net_return"], annualization, risk_free_rate, daily["nav"],
                       initial_nav, daily["date"], baseline_date)
    metrics.update(_relative_metrics(daily, annualization, risk_free_rate, baseline_date))
    attribution_error = abs(nav - initial_nav - daily["trading_pnl"].sum() + daily["total_cost"].sum())
    tolerance = 1e-10 * max(initial_nav, float(daily["nav"].max()))
    checks = pd.DataFrame([
        {"check": "daily_nav_accounting", "error": accounting_error, "tolerance": tolerance, "passed": accounting_error <= tolerance, "detail": "NAV = prior NAV + old-position P&L - fees - slippage"},
        {"check": "daily_contributions_sum", "error": contribution_error, "tolerance": 1e-10, "passed": contribution_error <= 1e-10, "detail": "Instrument net contributions sum to portfolio net return"},
        {"check": "post_cost_target_exposure", "error": target_error, "tolerance": 1e-10, "passed": target_error <= 1e-10, "detail": "Targets use equity after transaction costs"},
        {"check": "post_cost_capital_budget", "error": capital_target_error, "tolerance": 1e-9, "passed": capital_target_error <= 1e-9, "detail": "Post-trade n+m_B*x_B=1 and saved m_B; intraperiod breaches are recorded separately"},
        {"check": "cumulative_pnl_accounting", "error": attribution_error, "tolerance": tolerance * len(daily), "passed": attribution_error <= tolerance * len(daily), "detail": "Final NAV - initial NAV = sum(P&L) - sum(costs)"},
    ])
    if not checks["passed"].all():
        raise BacktestError("Independent accounting checks failed")
    summary = {"status": "completed", "engine_version": ENGINE_VERSION, "model": "weighted_futures_price_proxy",
               "start_date": daily["date"].iloc[0], "end_date": daily["date"].iloc[-1], "requested_start_date": start,
               "requested_end_date": end, "initial_nav": initial_nav, "final_nav": nav, **metrics,
               "baseline_date": baseline_date, "baseline_date_policy": baseline_policy,
               "baseline_date_is_observed": bool(first_position),
               **_turnover_metrics(daily, annualization),
               "cumulative_cost_rate": float(daily["cumulative_cost_rate"].iloc[-1]),
               "average_holding_count": float(daily["holding_count"].mean()),
               "target_volatility": signals[0]["target_volatility"], "rebalance_count": len(trades),
               "unexecuted_signal_count": sum(r["backtest_action"] != "executed" for r in signal_rows),
               "total_trading_pnl": float(daily["trading_pnl"].sum()), "total_fees": float(daily["fee"].sum()),
               "total_slippage": float(daily["slippage"].sum()), "total_cost": float(daily["total_cost"].sum()),
               "total_turnover": float(daily["turnover"].sum()), "mean_gross_exposure": float(daily["gross_exposure"].mean()),
               "max_gross_exposure": float(daily["gross_exposure"].max()), "max_non_bond_exposure": float(daily["non_bond_exposure"].max()),
               "bond_leverage_cap": leverage_cap, "bond_margin_floor": margin_floor,
               "mean_bond_margin_rate": daily["bond_margin_rate"].mean(),
               "minimum_bond_margin_rate": pd.concat([daily["bond_margin_rate"], daily["bond_margin_rate_before_trade"]]).min(),
               "maximum_bond_leverage": pd.concat([daily["bond_leverage"], daily["bond_leverage_before_trade"]]).max(),
               "margin_breach_days": int(daily["margin_breach_observed"].sum()),
               "margin_breach_days_after_trade": int(daily["margin_breach"].sum()),
               "bond_capital_depleted_days": int((daily["bond_capital_depleted"] | daily["bond_capital_depleted_before_trade"]).sum()),
               "max_capital_shortfall": max(daily["capital_shortfall"].max(), daily["capital_shortfall_before_trade"].max()),
               "capital_constraint_status": "期间出现资金约束违例" if daily["margin_breach_observed"].any() else "收盘观察未发现越界",
               "target_achieved_count": sum(t["optimization_status"] == "target_achieved" for t in trades),
               "cap_limited_count": sum(t["optimization_status"] == "cap_limited" for t in trades),
               "fee_bps": fee_bps, "slippage_bps": slippage_bps, "annualization": int(annualization),
               "risk_free_rate_for_sharpe": risk_free_rate, "rolling_window": int(rolling_window)}
    return {"summary": summary, "daily_nav": daily, "instrument_daily": pd.DataFrame(instrument_rows),
            "rebalance_allocations": pd.DataFrame(allocation_rows),
            "group_daily": pd.DataFrame(group_rows), "rebalances": pd.DataFrame(trades), "signals": pd.DataFrame(signal_rows),
            "attribution": pd.DataFrame(attribution_rows),
            "annual_returns": _period_returns(daily, "Y", annualization, risk_free_rate, baseline_date, baseline_policy),
            "monthly_returns": _period_returns(daily, "M", annualization, risk_free_rate, baseline_date, baseline_policy), "checks": checks}
