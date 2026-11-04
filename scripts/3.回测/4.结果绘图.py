from __future__ import annotations

from contextlib import contextmanager
import os
from pathlib import Path
import tempfile

import numpy as np
import pandas as pd

SUFFIXES = ("cumulative_nav", "drawdown", "turnover", "transaction_cost", "exposure", "bond_margin")
COLORS = ("#1F4E79", "#C0504D", "#548235", "#8064A2", "#008B8B")


def figure_paths(output_dir, mode="next_day_close"):
    return [Path(output_dir) / f"{mode}_{suffix}.png" for suffix in SUFFIXES]




def _check_destination(path, overwrite):
    if not isinstance(overwrite, bool):
        raise TypeError("overwrite must be bool")
    if path.suffix.lower() != ".png":
        raise ValueError("Figure must use .png")
    for item in (path, *path.absolute().parents):
        if item.is_symlink():
            raise ValueError(f"Refusing symbolic-link output path: {item}")
    if path.exists() and (not overwrite or not path.is_file()):
        raise FileExistsError(f"Figure already exists or is not a file: {path}")


@contextmanager
def _plot_context():
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.dates as mdates
        from matplotlib.ticker import PercentFormatter
    except ImportError as exc:
        raise ImportError("Plotting requires matplotlib in the project environment") from exc
    with plt.style.context("default"), matplotlib.rc_context({
        "font.sans-serif": ["Arial Unicode MS", "PingFang SC", "Heiti SC", "SimHei", "DejaVu Sans"],
        "axes.unicode_minus": False,
    }):
        yield plt, mdates, PercentFormatter


def _style(axis, title):
    axis.set_title(title, loc="left", fontsize=13, fontweight="bold")
    axis.grid(axis="y", color="#D9D9D9", linewidth=0.6)
    axis.spines[["top", "right"]].set_visible(False)
    axis.legend(frameon=False)


def _date_axis(axis, mdates):
    locator = mdates.AutoDateLocator(minticks=5, maxticks=9)
    axis.xaxis.set_major_locator(locator)
    axis.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))


def _save(figure, path, plt, overwrite):
    path = Path(path)
    temporary = None
    try:
        _check_destination(path, overwrite)
        path.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix=".backtest-figure-", suffix=".png", dir=path.parent)
        os.close(fd)
        figure.savefig(temporary, format="png", dpi=180, bbox_inches="tight", facecolor="white")
        _check_destination(path, overwrite)
        if overwrite:
            os.replace(temporary, path)
        else:
            os.link(temporary, path)
    finally:
        if temporary is not None:
            Path(temporary).unlink(missing_ok=True)
        plt.close(figure)


def _exposure_figure(daily, summary, plt, mdates, PercentFormatter):
    dates = pd.to_datetime(daily["date"]).to_numpy()
    rebalanced = daily["rebalance_no"].notna().to_numpy()
    nav = daily["nav"].to_numpy(dtype=float)
    nav_before = daily["nav_before_cost"].to_numpy(dtype=float)
    gross = daily["gross_exposure"].to_numpy(dtype=float)
    non_bond = daily["non_bond_exposure"].to_numpy(dtype=float)
    bond = daily["bond_exposure"].to_numpy(dtype=float)
    capital = daily["bond_capital"].to_numpy(dtype=float) / nav

    # 直接使用旧持仓按当日收盘价估值、尚未扣除本次交易费用时的状态。
    bond_before = daily["bond_exposure_before_trade"].to_numpy(dtype=float)
    non_bond_before = (daily["capital_used_at_floor_before_trade"].to_numpy(dtype=float)
                       - summary["bond_margin_floor"] * bond_before)
    gross_before = non_bond_before + bond_before
    capital_before = daily["bond_capital_before_trade"].to_numpy(dtype=float) / nav_before

    def mark_rebalances(axis, prior, after, *, legend=False):
        available = rebalanced & np.isfinite(prior)
        axis.plot(dates[available], prior[available], linestyle="none", marker="o", markersize=3.3,
                  markerfacecolor="white", markeredgecolor="#999999", markeredgewidth=.8,
                  label="调仓前（仅调仓日）" if legend else "_nolegend_", zorder=4)
        available = rebalanced & np.isfinite(after)
        axis.plot(dates[available], after[available], linestyle="none", marker="o", markersize=2.8,
                  markerfacecolor="#222222", markeredgecolor="white", markeredgewidth=.3,
                  label="调仓后（仅调仓日）" if legend else "_nolegend_", zorder=5)

    figure, axes = plt.subplots(2, 1, figsize=(10.6, 7.6), sharex=True, facecolor="white")
    axis = axes[0]
    axis.plot(dates, gross, color=COLORS[0], linewidth=1.5, label="总名义敞口")
    axis.plot(dates, bond, color=COLORS[3], linewidth=1.3, label="国债名义敞口")
    mark_rebalances(axis, gross_before, gross, legend=True)
    mark_rebalances(axis, bond_before, bond)
    axis.set_ylabel("名义金额 / 净资产")
    axis.yaxis.set_major_formatter(PercentFormatter(1., decimals=0))
    axis.set_ylim(bottom=-.1)
    label = f"次日收盘价；目标波动率{summary['target_volatility'] * 100:g}%；国债上限{summary['bond_leverage_cap']:g}倍"
    _style(axis, f"组合名义敞口（{label}）")
    axis.legend(frameon=False, fontsize=9.3, loc="upper left", ncol=2, labelspacing=.3)

    axis = axes[1]
    axis.plot(dates, non_bond, color=COLORS[1], linewidth=1.3, label="非债资金占比", zorder=2)
    axis.plot(dates, capital, color=COLORS[2], linewidth=1.2, label="国债资金占比", zorder=2)
    axis.axhline(1., color="#999999", linewidth=.8, linestyle="--", zorder=1)
    axis.axhline(0., color="#333333", linewidth=1.1, linestyle=":",
                 label="国债资金耗尽线：0%", zorder=2)
    mark_rebalances(axis, non_bond_before, non_bond, legend=True)
    # 首次建仓前尚未分配国债资金，保留缺失值；资金耗尽后的负值照常展示。
    mark_rebalances(axis, capital_before, capital)
    axis.set_ylabel("占净资产比例")
    axis.yaxis.set_major_formatter(PercentFormatter(1., decimals=0))
    funding_values = np.r_[non_bond, capital, non_bond_before[rebalanced], capital_before[rebalanced]]
    funding_values = funding_values[np.isfinite(funding_values)]
    low = min(-.1, np.floor((float(funding_values.min()) - .02) / .1) * .1)
    high = np.ceil((max(1., float(funding_values.max())) + .1) / .2) * .2
    axis.set_ylim(low, high)
    axis.set_yticks(np.r_[low, np.arange(0., high + .001, .2)])
    _style(axis, "资金分配（红线与绿线合计100%）")
    _date_axis(axis, mdates)
    handles, labels = axis.get_legend_handles_labels()
    order = [0, 1, 3, 4, 2]
    axis.legend([handles[i] for i in order], [labels[i] for i in order],
                frameon=False, fontsize=9.3, loc="upper right", ncol=3, labelspacing=.3)
    for axis in axes:
        axis.tick_params(labelsize=10)
        axis.margins(x=.05)
    figure.subplots_adjust(left=.085, right=.985, top=.95, bottom=.075, hspace=.25)
    return figure


def _bond_funding_figure(daily, summary, plt, mdates, PercentFormatter):
    from matplotlib.ticker import FixedLocator, FuncFormatter, NullLocator

    dates = pd.to_datetime(daily["date"]).to_numpy()
    rebalanced = daily["rebalance_no"].notna().to_numpy()
    floor, cap = summary["bond_margin_floor"], summary["bond_leverage_cap"]
    margin = daily["bond_margin_rate"].to_numpy(dtype=float)
    margin_before = daily["bond_margin_rate_before_trade"].to_numpy(dtype=float)
    leverage = daily["bond_leverage"].to_numpy(dtype=float)
    leverage_before = daily["bond_leverage_before_trade"].to_numpy(dtype=float)
    # 资金耗尽时杠杆无定义，在对数轴上保留断点，不以零或上限替代。
    leverage = np.where(daily["bond_capital"].to_numpy(dtype=float) > 0, leverage, np.nan)
    leverage_before = np.where(daily["bond_capital_before_trade"].to_numpy(dtype=float) > 0,
                               leverage_before, np.nan)

    figure, axes = plt.subplots(2, 1, figsize=(10.6, 7.6), sharex=True, facecolor="white")
    series = (
        (axes[0], margin, margin_before, floor, "保证金率", f"调仓保证金率下限：{floor * 100:.3g}%"),
        (axes[1], leverage, leverage_before, cap, "国债杠杆", f"调仓国债杠杆上限：{cap:g} 倍"),
    )
    for axis, values, before, limit, name, limit_label in series:
        axis.plot(dates, values, color=COLORS[0], linewidth=1., label=f"每日收盘{name}", zorder=2)
        available = rebalanced & np.isfinite(before)
        axis.plot(dates[available], before[available], linestyle="none", marker="o", markersize=3.,
                  markerfacecolor="white", markeredgecolor="#999999", markeredgewidth=.7,
                  label=f"调仓前{name}（仅调仓日）", zorder=3)
        axis.plot(dates[rebalanced], values[rebalanced], linestyle="none", marker=".", markersize=3.4,
                  color="#222222", label=f"调仓后{name}（仅调仓日）", zorder=4)
        axis.axhline(limit, color=COLORS[1], linestyle="--", linewidth=1.3, label=limit_label, zorder=1)
    axes[0].axhline(0., color="#333333", linestyle=":", linewidth=1.2,
                    label="国债资金耗尽线：0%", zorder=2)
    axes[0].yaxis.set_major_formatter(PercentFormatter(1., decimals=1))
    axes[0].set_ylabel("保证金率")
    # 将零线、全部每日值和实际调仓前的值纳入范围，保留负保证金率。
    margin_values = np.r_[0., floor, margin, margin_before[rebalanced]]
    margin_values = margin_values[np.isfinite(margin_values)]
    padding = max(float(np.ptp(margin_values)) * .05, floor * .1)
    axes[0].set_ylim(float(margin_values.min()) - padding, float(margin_values.max()) + padding)

    axes[1].set_yscale("log")
    axes[1].set_ylabel("倍数（对数刻度）")
    leverage_values = np.r_[cap, leverage, leverage_before[rebalanced]]
    leverage_values = leverage_values[np.isfinite(leverage_values) & (leverage_values > 0)]
    low, high = float(leverage_values.min()), float(leverage_values.max())
    decade = 10. ** np.floor(np.log10(high))
    upper_tick = np.ceil(high / decade - 1e-10) * decade
    lower_limit, upper_limit = low * .8, upper_tick * 1.25
    ticks = [multiple * 10. ** power
             for power in range(int(np.floor(np.log10(lower_limit))), int(np.ceil(np.log10(upper_limit))) + 1)
             for multiple in (1., 2., 5.)]
    ticks = sorted({tick for tick in (*ticks, cap, upper_tick) if lower_limit <= tick <= upper_limit})
    axes[1].set_ylim(lower_limit, upper_limit)
    axes[1].yaxis.set_major_locator(FixedLocator(ticks))
    axes[1].yaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:,.0f}" if value >= 1 else f"{value:g}"))
    axes[1].yaxis.set_minor_locator(NullLocator())

    label = f"次日收盘价；目标波动率{summary['target_volatility'] * 100:g}%；国债上限{cap:g}倍"
    _style(axes[0], f"国债资金监控（{label}）")
    _style(axes[1], "国债杠杆（对数刻度；国债资金不为正时留空）")
    for axis in axes:
        axis.legend(loc="upper right", frameon=False, fontsize=9.3, handlelength=2.2, labelspacing=.45)
        axis.tick_params(labelsize=10)
        axis.margins(x=.05)
    _date_axis(axes[1], mdates)

    # 峰值覆盖每日收盘和实际调仓前两种状态，避免遗漏调仓前的极值。
    observed_leverage = np.r_[leverage, leverage_before[rebalanced]]
    observed_dates = np.concatenate((dates, dates[rebalanced]))
    if np.isfinite(observed_leverage).any():
        peak = int(np.nanargmax(observed_leverage))
        if observed_leverage[peak] > cap * 2:
            day, value = pd.Timestamp(observed_dates[peak]), observed_leverage[peak]
            on_right = day > pd.Timestamp(dates[0]) + (pd.Timestamp(dates[-1]) - pd.Timestamp(dates[0])) / 2
            axes[1].annotate(f"{day:%Y-%m-%d}：{value:,.2f} 倍", xy=(day, value),
                             xytext=(-14 if on_right else 14, -2), textcoords="offset points",
                             ha="right" if on_right else "left", va="top", fontsize=8.8, color="#333333",
                             arrowprops=dict(arrowstyle="-", color="#666666", linewidth=.7))

    breach = daily["margin_breach"].to_numpy(dtype=bool)
    depleted = daily["bond_capital_depleted"].to_numpy(dtype=bool)
    passed = int(np.sum(rebalanced & np.isfinite(margin) & ~breach & ~depleted))
    count = int(rebalanced.sum())
    trade_note = (f"{count:,} 次实际调仓后均满足下限" if passed == count and count else
                  f"实际调仓后达标 {passed:,}/{count:,} 次")
    figure.subplots_adjust(left=.085, right=.985, top=.94, bottom=.13, hspace=.24)
    figure.text(.085, .061, f"{trade_note}；每日收盘保证金率低于下限 {int(breach.sum()):,} 天。",
                fontsize=9, color="#444444")
    figure.text(.085, .032, "当前回测仅在计划调仓日调整仓位，期间越界仅记录，未自动补足资金或减仓。",
                fontsize=9, color="#444444")
    return figure


def save_figures(output_dir, result, *, mode="next_day_close", overwrite=False):
    paths = figure_paths(output_dir, mode)
    for path in paths:
        _check_destination(path, overwrite)
    daily, summary = result["daily_nav"], result["summary"]
    required = {"date", "nav", "drawdown", "one_way_turnover", "cumulative_cost_rate", "gross_exposure", "non_bond_exposure"}
    if not required.issubset(daily.columns) or daily.empty:
        raise ValueError("Backtest result lacks required plotting fields")
    values = daily[list(required - {"date"})].to_numpy(dtype=float)
    if not np.isfinite(values).all():
        raise ValueError("Cannot plot nonfinite backtest values")
    dates = list(pd.to_datetime(daily["date"]))
    baseline = pd.Timestamp(summary["baseline_date"])
    if baseline >= dates[0]:
        raise ValueError("Figure baseline must precede the first return observation")
    comparison_dates = [baseline, *dates]
    label = f"次日收盘价；目标波动率{summary['target_volatility'] * 100:g}%；国债上限{summary['bond_leverage_cap']:g}倍"
    color = COLORS[0]
    with _plot_context() as (plt, mdates, PercentFormatter):
        figure, axis = plt.subplots(figsize=(10.0, 4.8), facecolor="white")
        axis.plot(comparison_dates, [1., *(daily["nav"] / summary["initial_nav"])], color=color, linewidth=1.5, label="全天候策略")
        axis.plot(comparison_dates, [1., *daily["benchmark_nav"]], color="#777777", linewidth=1.2, label="沪深300全收益")
        axis.set_ylabel("累计净值")
        _style(axis, f"Plot 1: Cumulative Net Value ({label})")
        _date_axis(axis, mdates)
        _save(figure, paths[0], plt, overwrite)

        figure, axis = plt.subplots(figsize=(10.0, 4.8), facecolor="white")
        drawdowns = np.r_[0., daily["drawdown"].to_numpy()]
        axis.plot(comparison_dates, drawdowns, color="#2867C7", linewidth=1.5, label="策略回撤")
        axis.plot(comparison_dates, [0., *daily["benchmark_drawdown"]], color="#777777", linewidth=1., label="基准回撤")
        axis.fill_between(comparison_dates, drawdowns, 0., color="#2867C7", alpha=.12)
        axis.axhline(0., color="#A6A6A6", linewidth=.8, linestyle="--")
        axis.set_ylabel("回撤")
        axis.yaxis.set_major_formatter(PercentFormatter(1.0))
        _style(axis, f"Plot 2: Drawdown ({label})")
        _date_axis(axis, mdates)
        _save(figure, paths[1], plt, overwrite)

        figure, axis = plt.subplots(figsize=(10.0, 4.5), facecolor="white")
        axis.bar(range(len(daily)), daily["one_way_turnover"], width=1., color=color, label="组合每日换仓比率（单边）")
        axis.set_xlabel("交易日序号")
        axis.yaxis.set_major_formatter(PercentFormatter(1.0))
        _style(axis, f"Plot 3: Rebalance Turnover ({label})")
        _save(figure, paths[2], plt, overwrite)

        figure, axis = plt.subplots(figsize=(10.0, 4.5), facecolor="white")
        axis.plot(dates, daily["cumulative_cost_rate"], color=color, linewidth=1.5, label="累计交易成本率")
        axis.yaxis.set_major_formatter(PercentFormatter(1.0))
        if not np.any(daily["cumulative_cost_rate"]):
            axis.set_ylim(-.0002, .01)
        _style(axis, f"Cumulative Transaction Cost Rate ({label})")
        _date_axis(axis, mdates)
        _save(figure, paths[3], plt, overwrite)

        figure = _exposure_figure(daily, summary, plt, mdates, PercentFormatter)
        _save(figure, paths[4], plt, overwrite)
        figure = _bond_funding_figure(daily, summary, plt, mdates, PercentFormatter)
        _save(figure, paths[5], plt, overwrite)
    return paths


