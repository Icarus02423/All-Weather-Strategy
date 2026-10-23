"""按已有调仓安排，估计股票、黄金、其他商品、国债的历史风险。

Python 3.9+，依赖 numpy、pandas；load_schedule 另需 openpyxl。
行情与调仓表各读一次，再逐期调用 estimate_risk。导入时不读取文件；
直接执行时只打印指定期的风险估计，不下载行情、不保存或修改数据。

口径：L 个简单日收益需要 L+1 个价格；整个窗口固定使用当期品种及
类内等分比例。年化协方差 = 252 × 样本协方差（ddof=1）。
"""
from __future__ import annotations

import argparse
import csv
import math
import re
from collections.abc import Mapping
from fractions import Fraction
from numbers import Integral, Real
from pathlib import Path
from typing import Any, Optional, Union

import numpy as np
import pandas as pd


ASSET_CLASSES = ("股票", "黄金", "其他商品", "国债")
TRADING_DAYS_PER_YEAR = 252
PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PRICE_PATH = PROJECT_ROOT / "data" / "futures" / "close_price.csv"
DEFAULT_SCHEDULE_PATH = PROJECT_ROOT / "data" / "prep" / "期货回测安排.xlsx"
PathLike = Union[str, Path]


def _as_date(value: Any, field: str) -> pd.Timestamp:
    if isinstance(value, str) and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        raise ValueError(f"{field}须使用 YYYY-MM-DD 格式。")
    if value is None or isinstance(value, (Real, bool)):
        raise ValueError(f"{field}不是有效日期。")
    try:
        result = pd.Timestamp(value)
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError(f"{field}不是有效日期：{value!r}。") from exc
    if pd.isna(result) or result.tzinfo is not None or result != result.normalize():
        raise ValueError(f"{field}须为不含时区及日内时间的有效日期。")
    return result


def _as_lookback(value: Any) -> int:
    if isinstance(value, bool) or not isinstance(value, Integral) or value < 2:
        raise ValueError("回看日数须为不小于2的整数，表示收益观测数。")
    return int(value)


def _validate_price_table(prices: pd.DataFrame) -> None:
    if not isinstance(prices, pd.DataFrame) or prices.empty:
        raise ValueError("价格表须为非空 pandas.DataFrame。")
    index = prices.index
    if not isinstance(index, pd.DatetimeIndex):
        raise ValueError("价格表须使用 DatetimeIndex 日期索引。")
    if index.hasnans or index.tz is not None or not index.equals(index.normalize()):
        raise ValueError("价格日期须有效，且不含时区及日内时间。")
    if not index.is_unique or not index.is_monotonic_increasing:
        raise ValueError("价格日期须唯一且按时间升序排列。")
    if not prices.columns.is_unique or any(
        not isinstance(code, str) or not code or code != code.strip()
        for code in prices.columns
    ):
        raise ValueError("价格列须为唯一、非空且无首尾空格的品种代码。")


def load_prices(path: PathLike = DEFAULT_PRICE_PATH) -> pd.DataFrame:
    """读取收盘价宽表，返回日期索引×品种代码的数值表。

    保留上市前的空白价格。只在 estimate_risk 中检查实际使用窗口的
    缺失值和非正价格，不对整张行情表 dropna 或填充。
    """
    path = Path(path)
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.reader(stream)
        header = next(reader, [])
        if len(header) < 2 or header[0] != "日期":
            raise ValueError("行情CSV首列须为“日期”，后续各列须为品种代码。")
        if len(set(header)) != len(header) or any(not x or x != x.strip() for x in header):
            raise ValueError("行情CSV列名存在重复、空白或首尾空格。")
        rows = []
        for row_number, row in enumerate(reader, start=2):
            if not row:
                continue
            if len(row) != len(header):
                raise ValueError(f"行情CSV第{row_number}行的列数与表头不一致。")
            rows.append(row)
    frame = pd.DataFrame(rows, columns=header)
    dates = [_as_date(value, "价格日期") for value in frame.pop("日期")]
    frame.index = pd.DatetimeIndex(dates, name="日期")
    for code in frame.columns:
        try:
            frame[code] = pd.to_numeric(frame[code].replace("", np.nan), errors="raise")
        except (ValueError, TypeError) as exc:
            raise ValueError(f"行情CSV的{code}列含非数字价格。") from exc
    _validate_price_table(frame)
    return frame.astype(float)


def _read_table(worksheet: Any, headers: tuple[str, ...]) -> list[tuple[int, dict]]:
    """按列名定位现有表，忽略表前说明及表后单列备注，不依赖行号。"""
    positions = None
    records = []
    ended = False
    for row_number, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
        if positions is None:
            if all(name in row for name in headers):
                if any(row.count(name) != 1 for name in headers):
                    raise ValueError(f"{worksheet.title}的表头存在重复列名。")
                positions = {name: row.index(name) for name in headers}
            continue
        record = {name: row[i] if i < len(row) else None for name, i in positions.items()}
        if all(value is None or value == "" for value in row):
            ended = True
            continue
        if ended:
            if any(value is not None and value != "" for value in row[1:]):
                raise ValueError(f"{worksheet.title}第{row_number}行在空行后仍有多列数据。")
            continue
        records.append((row_number, record))
    if positions is None or not records:
        raise ValueError(f"{worksheet.title}缺少所需表头或数据。")
    return records


def _validate_allocations(allocations: Mapping) -> dict[str, dict[str, float]]:
    if not isinstance(allocations, Mapping) or set(allocations) != set(ASSET_CLASSES):
        raise ValueError(f"类内分配须恰好包含四类：{ASSET_CLASSES}。")
    result = {}
    seen = set()
    for asset in ASSET_CLASSES:
        members = allocations[asset]
        if not isinstance(members, Mapping) or not members:
            raise ValueError(f"{asset}至少须有一个参与品种。")
        weights = {}
        for code, weight in members.items():
            if not isinstance(code, str) or not code or code != code.strip():
                raise ValueError(f"{asset}含无效品种代码。")
            if code in seen:
                raise ValueError(f"品种{code}在不同类别中重复出现。")
            if isinstance(weight, bool) or not isinstance(weight, Real):
                raise ValueError(f"{code}的类内比例须为数值。")
            value = float(weight)
            if not math.isfinite(value) or value <= 0:
                raise ValueError(f"{code}的类内比例须为有限正数。")
            weights[code] = value
            seen.add(code)
        equal_weight = 1.0 / len(weights)
        if abs(math.fsum(weights.values()) - 1.0) > 1e-12 or any(
            abs(weight - equal_weight) > 1e-12 for weight in weights.values()
        ):
            raise ValueError(f"{asset}须按品种数等分，类内比例之和须为1。")
        result[asset] = weights
    return result


def load_schedule(path: PathLike = DEFAULT_SCHEDULE_PATH) -> list[dict[str, Any]]:
    """只读既有“品种数据”“调仓安排”，返回每期可直接传入计算的安排。

    品种名称通过第一张表映射到完整行情代码；分数使用 Fraction 解析，
    不执行单元格文本。不重新推导上市资格、调仓日期或纳入日期。
    执行日为空的末期信号仍保留，可估计风险，是否交易由回测模块决定。
    """
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise ImportError("读取调仓xlsx需要openpyxl，请在当前Python环境安装。") from exc
    workbook = load_workbook(Path(path), read_only=True, data_only=True)
    try:
        if not {"品种数据", "调仓安排"}.issubset(workbook.sheetnames):
            raise ValueError("xlsx须包含“品种数据”和“调仓安排”两张表。")
        catalog = {}
        known_codes = set()
        for row_number, row in _read_table(
            workbook["品种数据"], ("资产类别", "品种名称", "代码")
        ):
            asset, name, code = (row[key] for key in ("资产类别", "品种名称", "代码"))
            if asset not in ASSET_CLASSES or not all(
                isinstance(value, str) and value.strip() == value and value
                for value in (name, code)
            ):
                raise ValueError(f"品种数据第{row_number}行的类别、名称或代码无效。")
            if (asset, name) in catalog or code in known_codes:
                raise ValueError(f"品种数据第{row_number}行的名称或代码重复。")
            catalog[(asset, name)] = code
            known_codes.add(code)
        allocation_headers = tuple(f"{asset}类内分配" for asset in ASSET_CLASSES)
        headers = ("回看日数", "调仓频率", "信号日", "执行日", *allocation_headers, "状态")
        schedule = []
        seen = set()
        for row_number, row in _read_table(workbook["调仓安排"], headers):
            try:
                lookback = _as_lookback(row["回看日数"])
                signal_date = _as_date(row["信号日"], "信号日")
                frequency = row["调仓频率"]
                if frequency not in ("每月", "每季度", "每半年"):
                    raise ValueError(f"未知调仓频率：{frequency!r}。")
                raw_execution = row["执行日"]
                execution_date = None if raw_execution in (None, "") else _as_date(raw_execution, "执行日")
                if execution_date is not None and execution_date <= signal_date:
                    raise ValueError("执行日须晚于信号日。")
                allocations = {}
                for asset, column in zip(ASSET_CLASSES, allocation_headers):
                    text = row[column]
                    if not isinstance(text, str) or not text.strip():
                        raise ValueError(f"{asset}的类内分配为空或不是文本。")
                    members = {}
                    for token in re.split(r"[；;]", text.strip()):
                        name, separator, proportion = token.replace("：", ":").partition(":")
                        key = (asset, name.strip())
                        if not separator or key not in catalog:
                            raise ValueError(f"{asset}含无法识别的品种分配：{token!r}。")
                        code = catalog[key]
                        if code in members:
                            raise ValueError(f"{code}在类内分配中重复出现。")
                        members[code] = float(Fraction(proportion.strip()))
                    allocations[asset] = members
                allocations = _validate_allocations(allocations)
                key = (lookback, frequency, signal_date)
                if key in seen:
                    raise ValueError("同一回看窗口、频率和信号日存在重复安排。")
                seen.add(key)
                schedule.append({
                    "signal_date": signal_date, "lookback": lookback,
                    "frequency": frequency, "execution_date": execution_date,
                    "allocations": allocations, "status": row["状态"],
                    "source_row": row_number,
                })
            except (ValueError, TypeError, ZeroDivisionError, OverflowError) as exc:
                raise ValueError(f"调仓安排第{row_number}行：{exc}") from exc
        return schedule
    finally:
        workbook.close()


def _get_returns(
    prices: pd.DataFrame, signal_date: pd.Timestamp, lookback: int,
    allocations: dict[str, dict[str, float]],
) -> tuple[pd.DataFrame, pd.Timestamp, pd.Timestamp]:
    _validate_price_table(prices)
    codes = [code for asset in ASSET_CLASSES for code in allocations[asset]]
    missing_codes = [code for code in codes if code not in prices.columns]
    if missing_codes:
        raise ValueError(f"价格表缺少参与品种：{missing_codes}。")
    if signal_date not in prices.index:
        raise ValueError("信号日不在行情交易日期中，不能自动退回其他日期。")
    end = prices.index.get_loc(signal_date)
    if end < lookback:
        raise ValueError(f"需要{lookback + 1}个价格，信号日及此前仅有{end + 1}行。")
    # 先固定同一段交易日期，再检查缺失；不能dropna后向前凑足观测数。
    window = prices.loc[:, codes].iloc[end - lookback:end + 1]
    if any(pd.api.types.is_bool_dtype(window[code]) for code in codes):
        raise ValueError("参与品种的价格不能是布尔值。")
    try:
        values = window.to_numpy(dtype=float, copy=True)
    except (ValueError, TypeError) as exc:
        raise ValueError("参与品种的价格含非数字内容。") from exc
    bad = np.argwhere(~np.isfinite(values) | (values <= 0))
    if len(bad):
        details = "、".join(f"{window.index[i]:%Y-%m-%d}/{codes[j]}" for i, j in bad[:5])
        raise ValueError(f"窗口内价格缺失、非有限或非正：{details}。")
    try:
        with np.errstate(over="raise", divide="raise", invalid="raise"):
            values = values[1:] / values[:-1] - 1.0
    except FloatingPointError as exc:
        raise ValueError("价格比值产生非有限收益率。") from exc
    if not np.isfinite(values).all():
        raise ValueError("收益率包含非有限值。")
    returns = pd.DataFrame(values, index=window.index[1:], columns=codes)
    return returns, window.index[0], window.index[-1]


def _get_class_returns(
    returns: pd.DataFrame, allocations: dict[str, dict[str, float]],
) -> pd.DataFrame:
    # 每一天使用当期固定的名义敞口比例，不拼接历史上不同名单的类别收益。
    columns = {}
    for asset in ASSET_CLASSES:
        weights = pd.Series(allocations[asset], dtype=float)
        columns[asset] = returns.loc[:, weights.index].dot(weights)
    return pd.DataFrame(columns, index=returns.index, columns=list(ASSET_CLASSES))


def _estimate_covariance(class_returns: pd.DataFrame) -> tuple[pd.DataFrame, pd.Series, dict]:
    try:
        with np.errstate(over="raise", invalid="raise"):
            matrix = np.cov(class_returns.to_numpy(dtype=float), rowvar=False, ddof=1)
            matrix = matrix * TRADING_DAYS_PER_YEAR
    except FloatingPointError as exc:
        raise ValueError("协方差计算发生数值溢出。") from exc
    if matrix.shape != (4, 4) or not np.isfinite(matrix).all():
        raise ValueError("协方差须为有限的4×4矩阵。")
    scale = max(float(np.max(np.abs(matrix))), np.finfo(float).tiny)
    tolerance = 1e-12 * scale
    if np.max(np.abs(matrix - matrix.T)) > tolerance:
        raise ValueError("协方差矩阵不对称。")
    try:
        eigenvalues = np.linalg.eigvalsh(matrix)
        rank = int(np.linalg.matrix_rank(matrix))
        condition = float(np.linalg.cond(matrix)) if rank == 4 else float("inf")
    except np.linalg.LinAlgError as exc:
        raise ValueError("协方差矩阵的数值检查未能完成。") from exc
    if eigenvalues[0] < -tolerance:
        raise ValueError("协方差矩阵不满足半正定条件。")
    if np.any(np.diag(matrix) < 0) or matrix[3, 3] <= 0:
        raise ValueError("各类方差不能为负，国债方差须为正。")
    labels = list(ASSET_CLASSES)
    covariance = pd.DataFrame(matrix, index=labels, columns=labels)
    volatility = pd.Series(np.sqrt(np.diag(matrix)), index=labels, name="年化波动率")
    diagnostics = {
        "eigenvalues": eigenvalues.copy(), "rank": rank,
        "condition_number": condition, "psd_tolerance": tolerance,
        "annualization_factor": TRADING_DAYS_PER_YEAR,
    }
    return covariance, volatility, diagnostics


def estimate_risk(
    prices: pd.DataFrame, signal_date: Any, lookback: int,
    allocations: Mapping[str, Mapping[str, float]],
) -> dict[str, Any]:
    """估计一个信号日、一个窗口的风险；不修改输入，不保存文件。

    allocations格式：{类别: {完整行情代码: 类内比例}}；四类均须非空，
    且每类按1/n等分。频率、风险预算、波动率参数不参与本函数。

    返回带类别标签的covariance、annual_volatility和class_returns，
    以及日期、样本数量、allocations和diagnostics。协方差单位为年化
    收益率平方；所有收益观测完整共同对齐，不做成对缺失协方差估计。
    """
    context = f"信号日={signal_date}，回看={lookback}"
    try:
        signal = _as_date(signal_date, "信号日")
        days = _as_lookback(lookback)
        weights = _validate_allocations(allocations)
        returns, price_start, price_end = _get_returns(prices, signal, days, weights)
        class_returns = _get_class_returns(returns, weights)
        covariance, volatility, diagnostics = _estimate_covariance(class_returns)
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError(f"风险估计失败（{context}）：{exc}") from exc
    return {
        "covariance": covariance, "asset_classes": ASSET_CLASSES,
        "annual_volatility": volatility, "class_returns": class_returns,
        "allocations": weights, "signal_date": signal, "lookback": days,
        "price_start": price_start, "price_end": price_end,
        "return_start": returns.index[0], "return_end": returns.index[-1],
        "n_prices": days + 1, "n_returns": len(returns), "diagnostics": diagnostics,
    }


def _main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="只读现有行情及安排，打印单期风险估计。")
    parser.add_argument("--prices", type=Path, default=DEFAULT_PRICE_PATH)
    parser.add_argument("--schedule", type=Path, default=DEFAULT_SCHEDULE_PATH)
    parser.add_argument("--signal-date", default="2014-12-31")
    parser.add_argument("--lookback", type=int, default=63)
    parser.add_argument("--frequency", choices=("每月", "每季度", "每半年"), default="每月")
    args = parser.parse_args(argv)
    try:
        signal = _as_date(args.signal_date, "信号日")
        schedule = load_schedule(args.schedule)
        matches = [row for row in schedule if (
            row["signal_date"] == signal and row["lookback"] == args.lookback
            and row["frequency"] == args.frequency
        )]
        if len(matches) != 1:
            raise ValueError("现有调仓表中未找到唯一匹配的日期、窗口和频率。")
        result = estimate_risk(load_prices(args.prices), signal, args.lookback, matches[0]["allocations"])
    except (ValueError, OSError, ImportError) as exc:
        parser.exit(1, f"{exc}\n")
    print(f"信号日：{signal:%Y-%m-%d}；回看：{result['n_returns']}个日收益")
    print(f"价格窗口：{result['price_start']:%Y-%m-%d} 至 {result['price_end']:%Y-%m-%d}")
    print("年化协方差矩阵：")
    print(result["covariance"].to_string(float_format=lambda value: f"{value:.8f}"))
    print("年化波动率：")
    print(result["annual_volatility"].to_string(float_format=lambda value: f"{value:.4%}"))
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
