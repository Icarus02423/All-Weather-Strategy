"""在 PyCharm 中直接运行本文件，完成全部135组参数的风险估计、优化和回测。

沿用 outputs/risk_estimate、outputs/optimization、outputs/backtest 的目录及格式。
每次运行重新计算并更新对应结果；不下载行情，不改调仓安排，不清空其他目录。
逐方案依次执行三个步骤，任何步骤失败立即停止，保留已完成结果供检查。
导入本文件不会运行任务；所有子程序使用当前 Python 解释器。
"""
from __future__ import annotations

from datetime import date
import importlib.util
from itertools import product
import math
from pathlib import Path
import subprocess
import sys
from time import perf_counter


PROJECT_ROOT = Path(__file__).resolve().parent

# 统一传给各阶段，不依赖各个脚本里单独运行时的默认窗口和调仓频率。
LOOKBACK_PERIODS = ("quarterly", "semiannual", "annual")  # 63、126、252个交易日
REBALANCE_FREQUENCIES = ("monthly", "quarterly", "semiannual")
TARGET_VOLATILITIES = (0.05, 0.10, 0.15)
BOND_LEVERAGE_CAPS = (50, 40, 30, 20, 10)
START_SIGNAL_DATE = "2014-12-31"  # 包含2015年首次建仓所需信号
START_DATE = "2015-01-01"
END_DATE = "2026-09-30"
INITIAL_NAV = 1.0
FEE_BPS = 10.0  # 单边0.1%，沿用当前手续费及滑点的合并估计
SLIPPAGE_BPS = 0.0
TRADING_DAYS_PER_YEAR = 252
RISK_FREE_RATE = 0.0
ROLLING_WINDOW = 63

LOOKBACK_NAMES = {"quarterly": "一季度回看", "semiannual": "半年回看", "annual": "一年回看"}
FREQUENCY_NAMES = {"monthly": "月度调仓", "quarterly": "季度调仓", "semiannual": "半年调仓"}
SCRIPTS = {
    "风险估计": "scripts/2.组合优化/1.风险估计.py",
    "组合优化": "scripts/2.组合优化/2.逐期运行.py",
    "回测与输出": "scripts/3.回测/2.运行回测.py",
}


def _preflight():
    """在任何输出更新前检查必需文件、依赖及日期设置。"""
    required = [
        *SCRIPTS.values(),
        "scripts/2.组合优化/3.优化求解.py",
        "scripts/3.回测/1.回测计算.py",
        "scripts/3.回测/3.结果读写.py",
        "scripts/3.回测/4.结果绘图.py",
        "data/futures/close_price.csv",
        "data/prep/期货回测安排.xlsx",
        "data/benchmark/H00300.CSI.csv",
    ]
    missing = [str(PROJECT_ROOT / name) for name in required if not (PROJECT_ROOT / name).is_file()]
    if missing:
        raise FileNotFoundError("缺少必需文件：\n" + "\n".join(missing))
    missing_modules = [name for name in ("numpy", "pandas", "openpyxl", "matplotlib")
                       if importlib.util.find_spec(name) is None]
    if missing_modules:
        raise ImportError(f"当前解释器 {sys.executable} 缺少依赖：{', '.join(missing_modules)}")
    if not date.fromisoformat(START_SIGNAL_DATE) <= date.fromisoformat(START_DATE) <= date.fromisoformat(END_DATE):
        raise ValueError("日期必须满足：首个信号日 ≤ 回测开始日 ≤ 回测结束日")
    if (not LOOKBACK_PERIODS or not REBALANCE_FREQUENCIES
            or any(value not in LOOKBACK_NAMES for value in LOOKBACK_PERIODS)
            or any(value not in FREQUENCY_NAMES for value in REBALANCE_FREQUENCIES)
            or len(set(LOOKBACK_PERIODS)) != len(LOOKBACK_PERIODS)
            or len(set(REBALANCE_FREQUENCIES)) != len(REBALANCE_FREQUENCIES)):
        raise ValueError("回看窗口和调仓频率必须有效、非空且不重复")
    for label, values in (("目标波动率", TARGET_VOLATILITIES), ("国债杠杆上限", BOND_LEVERAGE_CAPS)):
        if (not values or any(isinstance(value, bool) or not isinstance(value, (int, float))
                              or not math.isfinite(value) or value <= 0 for value in values)
                or len(set(values)) != len(values)):
            raise ValueError(f"{label}必须为非空、不重复的有限正数")


def _steps(lookback, frequency):
    """只构建命令，不运行；各阶段使用相同输入和参数。"""
    common = [
        "--lookback-period", lookback, "--frequency", frequency,
        "--prices", str(PROJECT_ROOT / "data/futures/close_price.csv"),
        "--schedule", str(PROJECT_ROOT / "data/prep/期货回测安排.xlsx"),
    ]
    signal_dates = ["--start-signal-date", START_SIGNAL_DATE, "--end-signal-date", END_DATE]
    scenarios = ["--targets", *map(str, TARGET_VOLATILITIES),
                 "--bond-leverage-caps", *map(str, BOND_LEVERAGE_CAPS)]
    return [
        ("风险估计", [
            *common, *signal_dates,
            "--annualization-days", str(TRADING_DAYS_PER_YEAR),
            "--output-root", str(PROJECT_ROOT / "outputs/risk_estimate"),
        ]),
        ("组合优化", [
            *common, *signal_dates, *scenarios,
            "--project-root", str(PROJECT_ROOT),
            "--risk-root", str(PROJECT_ROOT / "outputs/risk_estimate"),
            "--output-root", str(PROJECT_ROOT / "outputs/optimization"),
            "--failure-policy", "stop", "--trace-level", "full", "--overwrite",
        ]),
        ("回测与输出", [
            *common, *scenarios,
            "--project-root", str(PROJECT_ROOT),
            "--start-date", START_DATE, "--end-date", END_DATE,
            "--benchmark", str(PROJECT_ROOT / "data/benchmark/H00300.CSI.csv"),
            "--optimization-root", str(PROJECT_ROOT / "outputs/optimization"),
            "--output-root", str(PROJECT_ROOT / "outputs/backtest"),
            "--initial-nav", str(INITIAL_NAV), "--fee-bps", str(FEE_BPS),
            "--slippage-bps", str(SLIPPAGE_BPS),
            "--annualization", str(TRADING_DAYS_PER_YEAR),
            "--risk-free-rate", str(RISK_FREE_RATE), "--rolling-window", str(ROLLING_WINDOW),
            "--overwrite",
        ]),
    ]


def main():
    try:
        _preflight()
    except (OSError, ValueError, ImportError) as exc:
        print(f"全量回测未启动：{exc}", file=sys.stderr, flush=True)
        return 1
    schemes = list(product(LOOKBACK_PERIODS, REBALANCE_FREQUENCIES))
    per_scheme = len(TARGET_VOLATILITIES) * len(BOND_LEVERAGE_CAPS)
    total = len(schemes) * per_scheme
    print(f"开始全量运行：{len(schemes)}种窗口/频率方案，每种{per_scheme}组，共{total}组回测。", flush=True)
    print(f"回测区间：{START_DATE} 至 {END_DATE}；解释器：{sys.executable}", flush=True)
    print("对应方案的风险估计、优化和回测结果将重新生成，沿用原目录和文件名。", flush=True)
    started = perf_counter()
    completed = 0
    for number, (lookback, frequency) in enumerate(schemes, 1):
        scheme = f"{LOOKBACK_NAMES[lookback]} × {FREQUENCY_NAMES[frequency]}"
        for step_number, (stage, arguments) in enumerate(_steps(lookback, frequency), 1):
            print(f"\n[{number}/{len(schemes)}] {scheme}；步骤{step_number}/3：{stage}", flush=True)
            command = [sys.executable, "-B", "-u", str(PROJECT_ROOT / SCRIPTS[stage]), *arguments]
            try:
                # 直接传递参数列表，不经shell；子进程结束后释放本方案的大量结果数据。
                subprocess.run(command, cwd=PROJECT_ROOT, check=True)
            except (subprocess.CalledProcessError, OSError) as exc:
                print(f"\n全量运行已停止：{scheme}／{stage}失败。\n"
                      f"已完整完成{completed}/{total}组回测；本方案可能留下部分输出。\n"
                      f"请查看上方子程序的具体期数或文件报错，修正后重新运行本入口。\n{exc}",
                      file=sys.stderr, flush=True)
                return 1
            except KeyboardInterrupt:
                print(f"\n用户中断：{scheme}／{stage}；已完整完成{completed}/{total}组。",
                      file=sys.stderr, flush=True)
                return 130
        completed += per_scheme
        print(f"本方案完成；累计完成{completed}/{total}组回测。", flush=True)
    minutes = (perf_counter() - started) / 60
    print(f"\n全部完成：{completed}组回测，耗时{minutes:.1f}分钟。", flush=True)
    print(f"回测输出：{PROJECT_ROOT / 'outputs/backtest'}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
