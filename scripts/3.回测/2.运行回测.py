from __future__ import annotations

import argparse
import csv
import hashlib
import importlib.util
from pathlib import Path
import sys

import numpy as np
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
LOOKBACK_PERIOD = "quarterly"
REBALANCE_FREQUENCY = "monthly"
START_DATE = "2015-01-01"
END_DATE = "2026-09-30"
TARGET_VOLATILITIES = (0.05, 0.10, 0.15)
BOND_LEVERAGE_CAPS = (50, 40, 30, 20, 10)
BENCHMARK_CODE = "H00300.CSI"
BENCHMARK_NAME = "沪深300全收益指数"
INITIAL_NAV = 1.0
FEE_BPS = 10.0  # 单边0.1%，按成交名义金额计；合并估算手续费和滑点。
SLIPPAGE_BPS = 0.0  # 默认不另加滑点，避免重复计费。
TRADING_DAYS_PER_YEAR = 252
RISK_FREE_RATE = 0.0  # 仅用于 Sharpe；账户现金利息为零。
ROLLING_WINDOW = 63
OVERWRITE_OUTPUTS = False
EXECUTION_MODE = "next_day_close"

WINDOWS = {"quarterly": 63, "semiannual": 126, "annual": 252}
PERIODS = {"quarterly": "Quarterly", "semiannual": "Semiannual", "annual": "Annual"}
FREQUENCIES = {"monthly": "Monthly", "quarterly": "Quarterly", "semiannual": "Semiannual"}


def _load_module(path, name):
    path = Path(path)
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load {path}")
    source = path.read_bytes()
    module = importlib.util.module_from_spec(spec)
    previous = sys.modules.get(name)
    old = sys.dont_write_bytecode
    sys.dont_write_bytecode = True
    sys.modules[name] = module
    try:
        exec(compile(source, str(path), "exec"), module.__dict__)
        module.__source_sha256__ = hashlib.sha256(source).hexdigest()
    except BaseException:
        if previous is None:
            sys.modules.pop(name, None)
        else:
            sys.modules[name] = previous
        raise
    finally:
        sys.dont_write_bytecode = old
    return module


def _digest(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _date(value, label):
    if isinstance(value, (bool, int, float, np.number)) or value is None:
        raise ValueError(f"Invalid {label}")
    value = pd.Timestamp(value)
    if pd.isna(value) or value.tzinfo is not None or value != value.normalize():
        raise ValueError(f"Invalid {label}")
    return value


def _target_folder(target):
    percent = target * 100
    label = f"{int(percent):02d}" if percent.is_integer() else format(percent, ".10g").replace(".", "p")
    return f"TargetVol_{label}pct"


def _leverage_folder(limit):
    return f"BondLeverageCap_{format(limit, '.10g').replace('.', 'p')}x"


def _load_benchmark(path):
    frame = pd.read_csv(path, encoding="utf-8-sig", dtype={"代码": str})
    if list(frame.columns) != ["日期", "代码", "名称", "收盘价"] or frame.empty:
        raise ValueError("Benchmark must have 日期/代码/名称/收盘价 columns")
    if not frame["代码"].eq(BENCHMARK_CODE).all():
        raise ValueError("Unexpected benchmark code")
    dates = pd.DatetimeIndex(pd.to_datetime(frame["日期"], format="%Y-%m-%d", errors="raise"))
    values = pd.to_numeric(frame["收盘价"], errors="raise").to_numpy(dtype=float)
    if (dates.hasnans or not dates.is_unique or not dates.is_monotonic_increasing
            or not np.isfinite(values).all() or np.any(values <= 0)):
        raise ValueError("Invalid benchmark dates or closes")
    return pd.Series(values, index=dates, name=BENCHMARK_CODE)


def _check_outputs(destinations, output_root, protected, overwrite):
    if not isinstance(overwrite, bool):
        raise ValueError("overwrite must be boolean")
    resolved_output = output_root.resolve()
    for item in protected:
        item = Path(item).resolve()
        if item == resolved_output or item in resolved_output.parents or resolved_output in item.parents:
            raise ValueError(f"Output root must not overlap protected input/code: {item}")
    for destination in destinations:
        if destination.is_symlink() or (destination.exists() and (not overwrite or not destination.is_file())):
            raise FileExistsError(f"Refusing to replace output: {destination}")
        for parent in destination.parents:
            if parent.is_symlink():
                raise ValueError(f"Output path contains a symbolic link: {parent}")


def _check_schedule_coverage(calendar, records, start, end, frequency):
    if not records:
        raise ValueError("No preparation schedule rows for the selected scheme")
    months = {"monthly": 1, "quarterly": 3, "semiannual": 6}[frequency]
    # 使用完整行情日历识别各周期最后一个交易日，不能先截断到回测结束日。
    # 半年按自然年划分为1—6月、7—12月。
    period_closes = {}
    for day in calendar:
        period_closes[(day.year, (day.month - 1) // months)] = day
    final_period = (calendar[-1].year, (calendar[-1].month - 1) // months)
    actual = {record["signal_date"] for record in records}
    coverage_start = max(start, min(actual))
    missing = []
    for period, signal in period_closes.items():
        if not coverage_start <= signal <= end:
            continue
        year, part = period
        period_end = pd.Timestamp(year=year, month=(part + 1) * months, day=1) + pd.offsets.MonthEnd(0)
        # 后续周期已有行情，或已到自然期末，才能确认本周期已经结束。
        # 例如行情截至9月30日时，7—12月这半年的调仓尚未到期。
        completed = period != final_period or period_end <= end
        if completed and signal not in actual:
            missing.append(signal)
    if missing:
        dates = ", ".join(day.strftime("%Y-%m-%d") for day in missing[:5])
        raise ValueError(f"Preparation schedule is missing {frequency} signals: {dates}")


def run_backtests(*, project_root=PROJECT_ROOT, lookback_period=LOOKBACK_PERIOD,
                  rebalance_frequency=REBALANCE_FREQUENCY, start_date=START_DATE, end_date=END_DATE,
                  targets=TARGET_VOLATILITIES, leverage_caps=BOND_LEVERAGE_CAPS, initial_nav=INITIAL_NAV, fee_bps=FEE_BPS,
                  slippage_bps=SLIPPAGE_BPS, annualization=TRADING_DAYS_PER_YEAR,
                  risk_free_rate=RISK_FREE_RATE, rolling_window=ROLLING_WINDOW,
                  prices_path=None, schedule_path=None, benchmark_path=None, optimization_root=None,
                  output_root=None, overwrite=OVERWRITE_OUTPUTS):
    runner_hash = globals().get("__source_sha256__") or _digest(Path(__file__))
    root = Path(project_root).resolve()
    if lookback_period not in WINDOWS or rebalance_frequency not in FREQUENCIES:
        raise ValueError("Unknown lookback period or rebalance frequency")
    targets = tuple(targets)
    if (not targets or any(isinstance(t, (str, bool)) or not isinstance(t, (int, float, np.number))
                           or not np.isfinite(t) or t <= 0 for t in targets)):
        raise ValueError("Targets must be finite positive numbers")
    targets = tuple(float(t) for t in targets)
    if len(set(targets)) != len(targets) or len({_target_folder(t) for t in targets}) != len(targets):
        raise ValueError("Targets and output folder names must be distinct")
    leverage_caps = tuple(leverage_caps)
    if (not leverage_caps or any(isinstance(v, (str, bool)) or not isinstance(v, (int, float, np.number))
                                or not np.isfinite(v) or v <= 0 for v in leverage_caps)):
        raise ValueError("Bond leverage caps must be finite positive numbers")
    leverage_caps = tuple(float(v) for v in leverage_caps)
    if (len(set(leverage_caps)) != len(leverage_caps)
            or len({_leverage_folder(v) for v in leverage_caps}) != len(leverage_caps)
            or any(not np.isfinite(1. / v) or 1. / v <= 0 for v in leverage_caps)):
        raise ValueError("Invalid or duplicate bond leverage caps")
    prices_path = Path(prices_path) if prices_path is not None else root / "data/futures/close_price.csv"
    schedule_path = Path(schedule_path) if schedule_path is not None else root / "data/prep/期货回测安排.xlsx"
    benchmark_path = Path(benchmark_path) if benchmark_path is not None else root / "data/benchmark/H00300.CSI.csv"
    optimization_root = Path(optimization_root) if optimization_root is not None else root / "outputs/optimization"
    output_root = Path(output_root) if output_root is not None else root / "outputs/backtest"
    folder = f"Lookback_{PERIODS[lookback_period]}_Rebalance_{FREQUENCIES[rebalance_frequency]}"
    scenarios = [(target, limit) for target in targets for limit in leverage_caps]
    destinations = {(target, limit): output_root / folder / _target_folder(target) / _leverage_folder(limit) / EXECUTION_MODE
                    for target, limit in scenarios}
    comparison_path = output_root / folder / "parameter_comparison.xlsx"
    risk_path = root / "scripts/2.组合优化/1.风险估计.py"
    engine_path = Path(__file__).with_name("1.回测计算.py")
    io_path = Path(__file__).with_name("3.结果读写.py")
    plot_path = Path(__file__).with_name("4.结果绘图.py")
    risk = _load_module(risk_path, "_backtest_existing_risk_reader")
    engine = _load_module(engine_path, "_backtest_fixed_unit_engine")
    io = _load_module(io_path, "_backtest_result_io")
    plots = _load_module(plot_path, "_backtest_result_figures")
    output_files = [comparison_path, *plots.comparison_figure_paths(output_root / folder, targets)]
    for directory in destinations.values():
        output_files.extend(io.report_paths(directory, EXECUTION_MODE))
        output_files.extend(plots.figure_paths(directory, EXECUTION_MODE))
    _check_outputs(output_files, output_root,
                   [root / "scripts", root / "data", root / ".git", optimization_root, prices_path, schedule_path, benchmark_path], overwrite)
    price_hash, schedule_hash = _digest(prices_path), _digest(schedule_path)
    prices = risk.load_prices(prices_path)
    schedule = risk.load_schedule(schedule_path)
    benchmark_hash = _digest(benchmark_path)
    benchmark = _load_benchmark(benchmark_path)
    if _digest(benchmark_path) != benchmark_hash:
        raise ValueError("Benchmark changed while loading")
    if _digest(prices_path) != price_hash or _digest(schedule_path) != schedule_hash:
        raise ValueError("Input changed while loading")
    start = _date(start_date, "start_date") if start_date is not None else _date(START_DATE, "start_date")
    end = _date(end_date, "end_date") if end_date is not None else prices.index[-1]
    if start > end or start < prices.index[0] or end > prices.index[-1]:
        raise ValueError("Requested backtest dates are outside the available price range")
    planned = sorted([record for record in schedule
                       if record["lookback"] == WINDOWS[lookback_period] and record["frequency"] == rebalance_frequency
                       and record["signal_date"] <= end
                       and ((record["execution_date"] is not None and record["execution_date"] >= start)
                            or (record["execution_date"] is None and record["signal_date"] >= start))],
                      key=lambda record: record["signal_date"])
    for record in planned:
        signal_date, execution_date = record["signal_date"], record["execution_date"]
        if signal_date not in prices.index:
            raise ValueError("Scheduled signal is absent from the price calendar")
        position = prices.index.get_loc(signal_date)
        if execution_date is None:
            if position != len(prices.index) - 1:
                raise ValueError("Only the final observed signal may have no execution date")
        elif position + 1 >= len(prices.index) or execution_date != prices.index[position + 1]:
            raise ValueError("Scheduled execution must be the next observed date after signal")
    selected = [r for r in planned if r["execution_date"] is not None and r["execution_date"] <= end]
    pending = [r for r in planned if r["execution_date"] is None or r["execution_date"] > end]
    if not selected:
        raise ValueError("No executable preparation schedule rows match the requested settings")
    # 只要求截至结束日已经到期的信号齐全；最后一次调仓后可以继续持仓。
    matching_schedule = [r for r in schedule if r["lookback"] == WINDOWS[lookback_period] and r["frequency"] == rebalance_frequency]
    _check_schedule_coverage(prices.index, matching_schedule, start, end, rebalance_frequency)
    instrument_names, manifest_sources = {}, []
    manifest_path = root / "data/futures/download_manifest.csv"
    if manifest_path.is_file():
        manifest_hash = _digest(manifest_path)
        with manifest_path.open("r", encoding="utf-8-sig", newline="") as stream:
            for row in csv.DictReader(stream):
                code, name = row.get("代码"), row.get("名称")
                if not code or not name or code in instrument_names:
                    raise ValueError("Invalid or duplicate instrument names in download manifest")
                instrument_names[code] = name
        if _digest(manifest_path) != manifest_hash:
            raise ValueError("Instrument manifest changed while reading")
        manifest_sources.append({"kind": "instrument_manifest", "path": str(manifest_path.resolve()), "sha256": manifest_hash})
    signals_by_scenario, sources_by_scenario = {}, {}
    for target, limit in scenarios:
        signals, sources = [], []
        for index, record in enumerate(selected, 1):
            path = optimization_root / folder / _target_folder(target) / _leverage_folder(limit) / f"{record['rebalance_no']:03d}.xlsx"
            signal = io.read_optimization(path, record, target, prices=prices, leverage_cap=limit)
            signals.append(signal)
            sources.append({"kind": "optimization", "path": str(path.resolve()), "sha256": signal["source_sha256"],
                            "rebalance_no": record["rebalance_no"], "signal_date": record["signal_date"]})
            if index == 1 or index % 25 == 0 or index == len(selected):
                print(f"Validated target={target:.0%}, bond cap={limit:g}x: {index}/{len(selected)} optimization files", flush=True)
        signals_by_scenario[(target, limit)], sources_by_scenario[(target, limit)] = signals, sources
    inputs = {"lookback_period": lookback_period, "lookback_days": WINDOWS[lookback_period], "rebalance_frequency": rebalance_frequency,
              "start_date": start, "end_date": end, "initial_nav": initial_nav, "fee_bps": fee_bps, "slippage_bps": slippage_bps,
              "annualization": annualization, "risk_free_rate_for_sharpe": risk_free_rate, "rolling_window": rolling_window,
              "execution_timing": "Old holdings earn execution-day close-to-close P&L; rebalance at execution-day close; new holdings earn next observed-day P&L",
              "holding_policy": "Fixed synthetic price units between scheduled rebalances; weights drift; no daily risk-budget or cap rebalance",
              "initial_position_policy": "Start flat; positions before requested start are not inherited",
              "pending_signal_policy": "Signals with no in-range execution are recorded from the schedule; their optimization files are not required or read",
              "cost_policy": "Fee/slippage bps times absolute traded notional; target exposure denominator is post-cost equity",
              "turnover_policy": "Engine turnover is absolute buy/sell notional / pre-cost equity; ETF-format one-way turnover divides it by two; mean/annual turnover excludes first entry",
              "cumulative_cost_policy": "Daily cost rate=cost/pre-cost equity; cumulative cost rate=1-product(1-daily cost rate)",
              "benchmark_code": BENCHMARK_CODE, "benchmark_name": BENCHMARK_NAME, "benchmark_source": str(benchmark_path.resolve()),
              "benchmark_policy": "Same observed baseline and dates; no filling or sample truncation; relative NAV=strategy NAV/benchmark NAV",
              "margin_policy": "Close-price monitoring before and after scheduled rebalances; report breaches, no additional liquidation or trades",
              "leverage_caps": leverage_caps,
              "capacity_policy": "No turnover-volume input; capacity not estimated; portfolio exposure replaces capacity plot",
              "instrument_names": instrument_names,
              "cash_interest": 0., "terminal_policy": "Mark remaining holdings at last close; no forced liquidation cost",
              "data_model": "iFinD weighted prices; synthetic units and aggregate bond capital; no actual contract margins, integer lots or forced liquidation",
              "metric_policy": "Annualized return=(final/initial)^(annualization/N)-1; volatility=sample std(ddof=1)*sqrt(annualization); N includes cash/entry days",
              "attribution_policy": "Cumulative monetary P&L minus costs, divided by initial NAV; daily component returns are not separately compounded",
              "price_source": str(prices_path.resolve()), "schedule_source": str(schedule_path.resolve()),
              "optimization_root": str(optimization_root.resolve())}
    common_sources = [{"kind": kind, "path": str(Path(path).resolve()), "sha256": digest}
                      for kind, path, digest in (("prices", prices_path, price_hash), ("schedule", schedule_path, schedule_hash),
                                                ("engine", engine_path, engine.__source_sha256__), ("runner", Path(__file__), runner_hash),
                                                ("reader_exporter", io_path, io.__source_sha256__), ("risk_reader", risk_path, risk.__source_sha256__),
                                                ("plotter", plot_path, plots.__source_sha256__))]
    common_sources.extend(manifest_sources)
    common_sources.append({"kind": "benchmark", "path": str(benchmark_path.resolve()), "sha256": benchmark_hash})
    results = {}
    for (target, limit), signals in signals_by_scenario.items():
        key = (target, limit)
        results[key] = engine.backtest(prices, signals, benchmark_prices=benchmark, start_date=start, end_date=end, initial_nav=initial_nav,
                                         fee_bps=fee_bps, slippage_bps=slippage_bps, annualization=annualization,
                                         risk_free_rate=risk_free_rate, rolling_window=rolling_window)
        if pending:
            pending_rows = [{"rebalance_no": r["rebalance_no"], "signal_date": r["signal_date"],
                             "execution_date": r["execution_date"], "optimization_status": "not_loaded_not_required",
                             "backtest_action": "no_execution_date" if r["execution_date"] is None else "after_end_not_executed",
                             "target_volatility": target, "forecast_volatility": None,
                             "bond_leverage_cap": limit, "bond_margin_floor": 1. / limit, "target_bond_margin_rate": None,
                             "source_path": str(schedule_path.resolve()), "source_sha256": schedule_hash} for r in pending]
            results[key]["signals"] = pd.DataFrame.from_records(results[key]["signals"].to_dict("records") + pending_rows).sort_values("signal_date").reset_index(drop=True)
            results[key]["summary"]["unexecuted_signal_count"] += len(pending)
        print(f"Calculated target={target:.0%}, bond cap={limit:g}x: {results[key]['summary']['observations']} days, {results[key]['summary']['rebalance_count']} rebalances", flush=True)
    for key, result in results.items():
        target, limit = key
        io.save_backtest_reports(destinations[key], result, {**inputs, "target_volatility": target, "bond_leverage_cap": limit, "bond_margin_floor": 1. / limit},
                                 [*common_sources, *sources_by_scenario[key]], mode=EXECUTION_MODE, overwrite=overwrite)
        plots.save_figures(destinations[key], result, mode=EXECUTION_MODE, overwrite=overwrite)
        print(f"Saved 5 Excel reports + 6 PNG figures: {destinations[key]}", flush=True)
    io.save_comparison(comparison_path, results, inputs, overwrite=overwrite)
    plots.save_comparison_figures(output_root / folder, results, overwrite=overwrite)
    print(f"Saved {comparison_path}", flush=True)
    return {"status": "completed", "results": results, "paths": destinations,
            "comparison_path": comparison_path, "output_directory": output_root / folder}


def _main(argv=None):
    parser = argparse.ArgumentParser(description="Backtest existing optimized exposures using weighted futures close-price proxies.")
    parser.add_argument("--project-root", type=Path, default=PROJECT_ROOT)
    parser.add_argument("--lookback-period", choices=tuple(WINDOWS), default=LOOKBACK_PERIOD)
    parser.add_argument("--frequency", choices=tuple(FREQUENCIES), default=REBALANCE_FREQUENCY)
    parser.add_argument("--start-date", default=START_DATE)
    parser.add_argument("--end-date", default=END_DATE)
    parser.add_argument("--targets", nargs="+", type=float, default=TARGET_VOLATILITIES)
    parser.add_argument("--bond-leverage-caps", nargs="+", type=float, default=BOND_LEVERAGE_CAPS)
    parser.add_argument("--benchmark", type=Path)
    parser.add_argument("--initial-nav", type=float, default=INITIAL_NAV)
    parser.add_argument("--fee-bps", type=float, default=FEE_BPS)
    parser.add_argument("--slippage-bps", type=float, default=SLIPPAGE_BPS)
    parser.add_argument("--annualization", type=int, default=TRADING_DAYS_PER_YEAR)
    parser.add_argument("--risk-free-rate", type=float, default=RISK_FREE_RATE)
    parser.add_argument("--rolling-window", type=int, default=ROLLING_WINDOW)
    parser.add_argument("--prices", type=Path)
    parser.add_argument("--schedule", type=Path)
    parser.add_argument("--optimization-root", type=Path)
    parser.add_argument("--output-root", type=Path)
    parser.add_argument("--overwrite", action="store_true", default=OVERWRITE_OUTPUTS)
    args = parser.parse_args(argv)
    try:
        result = run_backtests(project_root=args.project_root, lookback_period=args.lookback_period,
                              rebalance_frequency=args.frequency, start_date=args.start_date, end_date=args.end_date,
                              targets=args.targets, leverage_caps=args.bond_leverage_caps, initial_nav=args.initial_nav, fee_bps=args.fee_bps,
                              slippage_bps=args.slippage_bps, annualization=args.annualization,
                              risk_free_rate=args.risk_free_rate, rolling_window=args.rolling_window,
                              prices_path=args.prices, schedule_path=args.schedule, benchmark_path=args.benchmark,
                              optimization_root=args.optimization_root, output_root=args.output_root, overwrite=args.overwrite)
    except (OSError, ValueError, TypeError, ImportError, OverflowError) as exc:
        parser.exit(1, f"Backtest failed: {exc}\n")
    print(f"Completed {len(result['results'])} parameter scenarios. Output: {result['output_directory']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
