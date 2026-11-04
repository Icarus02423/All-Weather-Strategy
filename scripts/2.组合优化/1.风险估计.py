"""按已有调仓安排，估计股票、黄金、其他商品、国债的历史风险。

Python 3.9+，依赖 numpy、pandas、openpyxl。
直接运行会计算顶部所选方案的完整历史，每次调仓单独保存一个英文 xlsx 文件。
文件名使用原调仓序号（001.xlsx、002.xlsx……）；重跑时更新该方案的结果。
导入时不读取或写入文件；不修改原始行情或调仓安排。

口径：L 个简单日收益需要 L+1 个价格；整个窗口固定使用当期品种及
类内等分比例。年化协方差 = 252 × 样本协方差（ddof=1）。
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import shutil
import tempfile
import math
import re
from collections.abc import Mapping
from fractions import Fraction
from datetime import date, datetime, timedelta, timezone
from numbers import Integral, Real
from pathlib import Path
from typing import Any, Optional, Union

import numpy as np
import pandas as pd


# 常用设置集中在这里。信号日是本次计算的数据截止日，不是历史行情起始日。
LOOKBACK_PERIOD = "quarterly"  # quarterly / semiannual / annual
REBALANCE_FREQUENCY = "quarterly"  # monthly / quarterly / semiannual
START_SIGNAL_DATE = "2014-12-31"  # 为2015年起的回测准备首期风险；每期自动向前回看。
END_SIGNAL_DATE = "2026-09-30"  # 信号日范围包含起止两端，筛选后保留原调仓序号。
TRADING_DAYS_PER_YEAR = 252  # 日协方差的年化系数，与回看周期分别设置。
PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PRICE_PATH = PROJECT_ROOT / "data" / "futures" / "close_price.csv"
DEFAULT_SCHEDULE_PATH = PROJECT_ROOT / "data" / "prep" / "期货回测安排.xlsx"
OUTPUT_ROOT = PROJECT_ROOT / "outputs" / "risk_estimate"

LOOKBACK_WINDOWS = {"quarterly": 63, "semiannual": 126, "annual": 252}
ASSET_CLASSES = ("equity", "gold", "other_commodities", "bonds")
INPUT_CLASS_NAMES = dict(zip(("股票", "黄金", "其他商品", "国债"), ASSET_CLASSES))
INPUT_FREQUENCIES = {"每月": "monthly", "每季度": "quarterly", "每半年": "semiannual"}
INPUT_STATUSES = {
    "首次建仓": "initial_allocation", "按规则调仓": "scheduled_rebalance",
    "样本结束，无后续执行日": "no_execution_date",
}
PathLike = Union[str, Path]


def _as_date(value: Any, field: str) -> pd.Timestamp:
    if isinstance(value, str) and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        raise ValueError(f"{field} must use YYYY-MM-DD.")
    if value is None or isinstance(value, (Real, bool)):
        raise ValueError(f"{field} is not a valid date.")
    try:
        result = pd.Timestamp(value)
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError(f"Invalid {field}: {value!r}.") from exc
    if pd.isna(result) or result.tzinfo is not None or result != result.normalize():
        raise ValueError(f"{field} must be a valid date without a time zone or time of day.")
    return result


def _as_lookback(value: Any) -> int:
    if isinstance(value, bool) or not isinstance(value, Integral) or value < 2:
        raise ValueError("lookback must be an integer of at least 2 daily return observations.")
    return int(value)


def _validate_price_table(prices: pd.DataFrame) -> None:
    if not isinstance(prices, pd.DataFrame) or prices.empty:
        raise ValueError("prices must be a nonempty pandas.DataFrame.")
    index = prices.index
    if not isinstance(index, pd.DatetimeIndex):
        raise ValueError("prices must have a DatetimeIndex.")
    if index.hasnans or index.tz is not None or not index.equals(index.normalize()):
        raise ValueError("Price dates must be valid dates without time zones or intraday times.")
    if not index.is_unique or not index.is_monotonic_increasing:
        raise ValueError("Price dates must be unique and sorted in ascending order.")
    if not prices.columns.is_unique or any(
        not isinstance(code, str) or not code or code != code.strip()
        for code in prices.columns
    ):
        raise ValueError("Price columns must be unique nonempty instrument codes without surrounding spaces.")


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
            raise ValueError("The price CSV must start with the date column followed by instrument codes.")
        if len(set(header)) != len(header) or any(not x or x != x.strip() for x in header):
            raise ValueError("Price CSV column names are duplicated, blank or contain surrounding spaces.")
        rows = []
        for row_number, row in enumerate(reader, start=2):
            if not row:
                continue
            if len(row) != len(header):
                raise ValueError(f"Price CSV row {row_number} has an unexpected number of fields.")
            rows.append(row)
    frame = pd.DataFrame(rows, columns=header)
    dates = [_as_date(value, "price_date") for value in frame.pop("日期")]
    frame.index = pd.DatetimeIndex(dates, name="date")
    for code in frame.columns:
        try:
            frame[code] = pd.to_numeric(frame[code].replace("", np.nan), errors="raise")
        except (ValueError, TypeError) as exc:
            raise ValueError(f"Price CSV column {code} contains nonnumeric prices.") from exc
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
                    raise ValueError("Worksheet headers contain duplicate column names.")
                positions = {name: row.index(name) for name in headers}
            continue
        record = {name: row[i] if i < len(row) else None for name, i in positions.items()}
        if all(value is None or value == "" for value in row):
            ended = True
            continue
        if ended:
            if any(value is not None and value != "" for value in row[1:]):
                raise ValueError(f"Worksheet row {row_number} has multi-column data after the table ended.")
            continue
        records.append((row_number, record))
    if positions is None or not records:
        raise ValueError("Worksheet is missing required headers or data.")
    return records


def _validate_allocations(allocations: Mapping) -> dict[str, dict[str, float]]:
    if not isinstance(allocations, Mapping) or set(allocations) != set(ASSET_CLASSES):
        raise ValueError(f"Allocations must contain exactly these classes: {ASSET_CLASSES}.")
    result = {}
    seen = set()
    for asset in ASSET_CLASSES:
        members = allocations[asset]
        if not isinstance(members, Mapping) or not members:
            raise ValueError(f"{asset} must contain at least one instrument.")
        weights = {}
        for code, weight in members.items():
            if not isinstance(code, str) or not code or code != code.strip():
                raise ValueError(f"{asset} contains an invalid instrument code.")
            if code in seen:
                raise ValueError(f"Instrument {code} appears in more than one asset class.")
            if isinstance(weight, bool) or not isinstance(weight, Real):
                raise ValueError(f"The within-class weight for {code} must be numeric.")
            value = float(weight)
            if not math.isfinite(value) or value <= 0:
                raise ValueError(f"The within-class weight for {code} must be finite and positive.")
            weights[code] = value
            seen.add(code)
        equal_weight = 1.0 / len(weights)
        if abs(math.fsum(weights.values()) - 1.0) > 1e-12 or any(
            abs(weight - equal_weight) > 1e-12 for weight in weights.values()
        ):
            raise ValueError(f"{asset} must use equal within-class weights summing to 1.")
        result[asset] = weights
    return result


def load_schedule(path: PathLike = DEFAULT_SCHEDULE_PATH) -> list[dict[str, Any]]:
    """读取现有中文准备表，返回英文标识及原调仓序号，不重算安排。"""
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise ImportError("Reading the schedule requires openpyxl: python -m pip install openpyxl") from exc
    workbook = load_workbook(Path(path), read_only=True, data_only=True)
    try:
        if not {"品种数据", "调仓安排"}.issubset(workbook.sheetnames):
            raise ValueError("The schedule workbook must contain instrument and rebalance worksheets.")
        catalog, known_codes = {}, set()
        for row_number, row in _read_table(workbook["品种数据"], ("资产类别", "品种名称", "代码")):
            input_asset, name, code = (row[key] for key in ("资产类别", "品种名称", "代码"))
            if input_asset not in INPUT_CLASS_NAMES or not all(
                isinstance(value, str) and value and value.strip() == value for value in (name, code)
            ):
                raise ValueError(f"Instrument row {row_number}: invalid class, name or code.")
            asset = INPUT_CLASS_NAMES[input_asset]
            if (asset, name) in catalog or code in known_codes:
                raise ValueError(f"Instrument row {row_number}: duplicate name or code.")
            catalog[(asset, name)] = code
            known_codes.add(code)
        allocation_headers = tuple(f"{asset}类内分配" for asset in INPUT_CLASS_NAMES)
        headers = ("调仓序号", "回看日数", "调仓频率", "信号日", "执行日", *allocation_headers, "状态")
        schedule, seen, groups = [], set(), {}
        for row_number, row in _read_table(workbook["调仓安排"], headers):
            try:
                number = row["调仓序号"]
                if isinstance(number, bool) or not isinstance(number, Integral) or number < 1:
                    raise ValueError("rebalance_no must be a positive integer.")
                lookback = _as_lookback(row["回看日数"])
                signal = _as_date(row["信号日"], "signal_date")
                if row["调仓频率"] not in INPUT_FREQUENCIES:
                    raise ValueError("Unknown rebalance frequency.")
                frequency = INPUT_FREQUENCIES[row["调仓频率"]]
                raw_execution = row["执行日"]
                execution = None if raw_execution in (None, "") else _as_date(raw_execution, "execution_date")
                if execution is not None and execution <= signal:
                    raise ValueError("execution_date must be after signal_date.")
                if row["状态"] not in INPUT_STATUSES:
                    raise ValueError("Unknown schedule status.")
                allocations = {}
                for asset, column in zip(ASSET_CLASSES, allocation_headers):
                    text = row[column]
                    if not isinstance(text, str) or not text.strip():
                        raise ValueError(f"Missing allocation text for {asset}.")
                    members = {}
                    for token in re.split(r"[；;]", text.strip()):
                        name, separator, proportion = token.replace("：", ":").partition(":")
                        key = (asset, name.strip())
                        if not separator or key not in catalog:
                            raise ValueError(f"Unrecognized allocation in {asset}.")
                        code = catalog[key]
                        if code in members:
                            raise ValueError(f"Duplicate instrument in allocation: {code}.")
                        members[code] = float(Fraction(proportion.strip()))
                    allocations[asset] = members
                allocations = _validate_allocations(allocations)
                key = (lookback, frequency, signal)
                if key in seen:
                    raise ValueError("Duplicate lookback, frequency and signal_date.")
                seen.add(key)
                record = {
                    "rebalance_no": int(number), "signal_date": signal,
                    "lookback": lookback, "frequency": frequency,
                    "execution_date": execution, "allocations": allocations,
                    "status": INPUT_STATUSES[row["状态"]], "source_row": row_number,
                }
                schedule.append(record)
                groups.setdefault((lookback, frequency), []).append(record)
            except (ValueError, TypeError, ZeroDivisionError, OverflowError) as exc:
                raise ValueError(f"Schedule row {row_number}: {exc}") from exc
        # 先校验完整安排的序号，再按用户日期范围筛选；不重新编号。
        for key, records in groups.items():
            ordered = sorted(records, key=lambda item: item["signal_date"])
            if [item["rebalance_no"] for item in ordered] != list(range(1, len(ordered) + 1)):
                raise ValueError(f"Schedule {key}: rebalance_no must follow signal dates consecutively from 1.")
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
        raise ValueError(f"Missing selected instruments in prices: {missing_codes}.")
    if signal_date not in prices.index:
        raise ValueError("signal_date is not an observed trading date; no date substitution is allowed.")
    end = prices.index.get_loc(signal_date)
    if end < lookback:
        raise ValueError(f"Need {lookback + 1} prices; only {end + 1} rows exist through signal_date.")
    # 先固定同一段交易日期，再检查缺失；不能dropna后向前凑足观测数。
    window = prices.loc[:, codes].iloc[end - lookback:end + 1]
    if any(pd.api.types.is_bool_dtype(window[code]) for code in codes):
        raise ValueError("Selected prices cannot be boolean values.")
    try:
        values = window.to_numpy(dtype=float, copy=True)
    except (ValueError, TypeError) as exc:
        raise ValueError("Selected prices contain nonnumeric values.") from exc
    bad = np.argwhere(~np.isfinite(values) | (values <= 0))
    if len(bad):
        details = ", ".join(f"{window.index[i]:%Y-%m-%d}/{codes[j]}" for i, j in bad[:5])
        raise ValueError(f"Missing, nonfinite or nonpositive window prices: {details}.")
    try:
        with np.errstate(over="raise", divide="raise", invalid="raise"):
            values = values[1:] / values[:-1] - 1.0
    except FloatingPointError as exc:
        raise ValueError("Price ratios produced nonfinite returns.") from exc
    if not np.isfinite(values).all():
        raise ValueError("Returns contain nonfinite values.")
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


def _estimate_covariance(class_returns: pd.DataFrame, annualization_days: int) -> tuple[pd.DataFrame, pd.Series, dict]:
    try:
        with np.errstate(over="raise", invalid="raise"):
            matrix = np.cov(class_returns.to_numpy(dtype=float), rowvar=False, ddof=1)
            matrix = matrix * annualization_days
    except FloatingPointError as exc:
        raise ValueError("Covariance calculation overflowed.") from exc
    if matrix.shape != (4, 4) or not np.isfinite(matrix).all():
        raise ValueError("Covariance must be a finite 4-by-4 matrix.")
    scale = max(float(np.max(np.abs(matrix))), np.finfo(float).tiny)
    tolerance = 1e-12 * scale
    if np.max(np.abs(matrix - matrix.T)) > tolerance:
        raise ValueError("Covariance is not symmetric.")
    try:
        eigenvalues = np.linalg.eigvalsh(matrix)
        rank = int(np.linalg.matrix_rank(matrix))
        condition = float(np.linalg.cond(matrix)) if rank == 4 else float("inf")
    except np.linalg.LinAlgError as exc:
        raise ValueError("Covariance diagnostics could not be completed.") from exc
    if eigenvalues[0] < -tolerance:
        raise ValueError("Covariance is not positive semidefinite within tolerance.")
    if np.any(np.diag(matrix) < 0) or matrix[3, 3] <= 0:
        raise ValueError("Variances must be nonnegative and bond variance must be positive.")
    labels = list(ASSET_CLASSES)
    covariance = pd.DataFrame(matrix, index=labels, columns=labels)
    volatility = pd.Series(np.sqrt(np.diag(matrix)), index=labels, name="annualized_volatility")
    diagnostics = {
        "eigenvalues": eigenvalues.copy(), "rank": rank,
        "condition_number": condition, "psd_tolerance": tolerance,
        "annualization_factor": annualization_days,
    }
    return covariance, volatility, diagnostics


def estimate_risk(
    prices: pd.DataFrame, signal_date: Any, lookback: int,
    allocations: Mapping[str, Mapping[str, float]], annualization_days: Optional[int] = None,
) -> dict[str, Any]:
    """估计一个信号日、一个窗口的风险；不修改输入，不保存文件。

    allocations格式：{类别: {完整行情代码: 类内比例}}；四类均须非空，
    且每类按1/n等分。频率、风险预算、波动率参数不参与本函数。

    返回带类别标签的covariance、annual_volatility和class_returns，
    以及未乘类内权重的instrument_returns和对应的instrument_covariance。
    同时返回日期、样本数量、allocations和diagnostics。协方差单位为年化
    收益率平方；所有收益观测完整共同对齐，不做成对缺失协方差估计。
    """
    context = f"signal_date={signal_date}, lookback={lookback}"
    try:
        signal = _as_date(signal_date, "signal_date")
        days = _as_lookback(lookback)
        weights = _validate_allocations(allocations)
        factor = _annualization_days(TRADING_DAYS_PER_YEAR if annualization_days is None else annualization_days)
        returns, price_start, price_end = _get_returns(prices, signal, days, weights)
        class_returns = _get_class_returns(returns, weights)
        covariance, volatility, diagnostics = _estimate_covariance(class_returns, factor)
        with np.errstate(over="raise", invalid="raise"):
            instrument_matrix = np.cov(returns.to_numpy(dtype=float), rowvar=False, ddof=1) * factor
        if instrument_matrix.shape != (len(returns.columns), len(returns.columns)) or not np.isfinite(instrument_matrix).all():
            raise ValueError("Instrument covariance must be a finite matrix covering the selected instruments.")
        instrument_covariance = pd.DataFrame(
            instrument_matrix, index=returns.columns, columns=returns.columns,
        )
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError(f"Risk estimation failed ({context}): {exc}") from exc
    return {
        "covariance": covariance, "asset_classes": ASSET_CLASSES,
        "annual_volatility": volatility, "class_returns": class_returns,
        "instrument_returns": returns, "instrument_covariance": instrument_covariance,
        "allocations": weights, "signal_date": signal, "lookback": days,
        "price_start": price_start, "price_end": price_end,
        "return_start": returns.index[0], "return_end": returns.index[-1],
        "n_prices": days + 1, "n_returns": len(returns), "diagnostics": diagnostics,
    }



def _workbook_cell_value(value: Any, sheet: str, column: str, row: int) -> Any:
    """Keep Excel-native values; reject nonfinite financial observations."""
    if value is None or value is pd.NA or value is pd.NaT:
        return None
    if isinstance(value, np.datetime64):
        value = pd.Timestamp(value)
    elif isinstance(value, np.generic):
        value = value.item()
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None
        if value.tzinfo is not None:
            raise ValueError(f"{sheet}.{column}, row {row}: dates must have no time zone.")
        return value.to_pydatetime()
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            raise ValueError(f"{sheet}.{column}, row {row}: dates must have no time zone.")
        return value
    if isinstance(value, date):
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, Integral):
        return int(value)
    if isinstance(value, Real):
        value = float(value)
        if not math.isfinite(value):
            if "condition" in column.lower() and math.isinf(value) and value > 0:
                return "unbounded"
            raise ValueError(f"{sheet}.{column}, row {row}: nonfinite numeric value.")
        return value
    if isinstance(value, str):
        return value
    raise TypeError(
        f"{sheet}.{column}, row {row}: unsupported cell type {type(value).__name__}."
    )


def _risk_column_format(sheet: str, column: str, value: Any) -> str:
    if isinstance(value, (datetime, date)):
        return "yyyy-mm-dd"
    if isinstance(value, bool):
        return "General"
    if isinstance(value, Integral):
        return "#,##0"
    if not isinstance(value, Real):
        return "General"
    name = column.lower()
    assets = {"equity", "gold", "other_commodities", "bonds"}
    if sheet == "Covariance" and name in assets:
        return "0.0000000000"
    if (
        name == "within_class_weight"
        or "volatility" in name
        or name.endswith("_vol")
        or (sheet == "ClassReturns" and name in assets)
        or (sheet == "InstrumentReturns" and name != "return_date")
    ):
        return "0.0000%"
    if "eigenvalue" in name or "tolerance" in name:
        return "0.0000000000E+00"
    return "0.0000000000"


def _validate_output_tables(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Validate the full batch before any per-period filtering or writing."""
    expected = ("Overview", "Covariance", "Allocations", "ClassReturns",
                "InstrumentCovariance", "InstrumentReturns")
    if set(tables) != set(expected):
        raise ValueError(f"Output tables must contain exactly {expected}.")
    identity_columns = (
        "rebalance_no", "signal_date", "lookback_period", "lookback_days",
        "rebalance_frequency",
    )
    for name in expected:
        frame = tables[name]
        if not isinstance(frame, pd.DataFrame) or frame.empty:
            raise ValueError(f"{name} must be a nonempty DataFrame.")
        if not frame.columns.is_unique or any(
            not isinstance(column, str) or not column for column in frame.columns
        ):
            raise ValueError(f"{name} must have unique, nonempty string column names.")
        if not set(identity_columns).issubset(frame.columns):
            raise ValueError(f"{name} is missing period identity columns.")
    overview = tables["Overview"]
    if overview["rebalance_no"].duplicated().any():
        raise ValueError("Overview contains duplicate rebalance numbers.")
    for number in overview["rebalance_no"]:
        if isinstance(number, bool) or not isinstance(number, Integral) or number < 1:
            raise ValueError("rebalance_no must be a positive integer.")
    period_numbers = set(overview["rebalance_no"])
    for name in expected[1:]:
        if set(tables[name]["rebalance_no"]) != period_numbers:
            raise ValueError(f"{name} must contain exactly the Overview rebalance numbers.")

    return overview


def _write_risk_workbook(tables: dict[str, pd.DataFrame], destination: Path) -> None:
    """Write one rebalance to its own workbook, using one plain worksheet.

    The worksheet is named Rebalance_<original schedule number, padded to 3>.
    Rows 3--11 hold core period parameters/sample information in A:B and D:E.
    AnnualCovariance starts at A14, with its header at A15:E15 and data in
    A16:E19. AnnualVolatility starts at G14, header G15:H15, data G16:H19.
    Allocations starts at A22 (header row 23). ClassReturns follows after one
    empty row. Diagnostics is at G22 beside the class covariance/volatility;
    data_version is in the top parameter area at G3:H3.
    InstrumentAnnualCovariance starts at J14; unweighted
    InstrumentReturns starts in column J alongside the existing ClassReturns.
    InstrumentAnnualVolatility is placed to the right of its covariance, with
    InstrumentDiagnostics below it. Existing covariance/return/allocation
    cells, their column widths and the original filter remain unchanged.
    Full run metadata is retained in workbook properties.
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font
        from openpyxl.utils import get_column_letter
    except ImportError as exc:
        raise ImportError(
            "Writing the output workbook requires openpyxl. "
            "Install it in the current environment: python -m pip install openpyxl"
        ) from exc

    overview = _validate_output_tables(tables)
    if len(overview) != 1:
        raise ValueError("Each workbook must contain exactly one rebalance.")
    expected = ("Overview", "Covariance", "Allocations", "ClassReturns",
                "InstrumentCovariance", "InstrumentReturns")
    identity_columns = (
        "rebalance_no", "signal_date", "lookback_period", "lookback_days",
        "rebalance_frequency",
    )

    workbook = Workbook(iso_dates=False)
    workbook.remove(workbook.active)
    metadata = overview.attrs.get("run_metadata", {})
    workbook.properties.description = json.dumps(metadata, ensure_ascii=True, default=str)
    font = Font(name="Carlito", size=11, color="000000", bold=False)
    alignment = Alignment(horizontal="center", vertical="center")
    left_fields = (
        "rebalance_no", "signal_date", "execution_date", "lookback_period",
        "lookback_days", "rebalance_frequency", "annualization_days",
        "covariance_ddof", "return_method",
    )
    right_fields = (
        "price_start", "price_end", "return_start", "return_end", "n_prices",
        "n_returns", "risk_status", "execution_status", "run_date",
    )
    volatility_fields = tuple(f"{asset}_volatility" for asset in ASSET_CLASSES)
    required_overview = set(left_fields + right_fields + volatility_fields + ("data_version",))
    if not required_overview.issubset(overview.columns):
        raise ValueError("Overview is missing required period or volatility fields.")

    def put(worksheet: Any, row: int, column: int, value: Any,
            section: str, field: str) -> None:
        value = _workbook_cell_value(value, section, field, row)
        cell = worksheet.cell(row, column, value)
        if isinstance(value, str):
            # All labels and source strings are literal text, never formulas.
            cell.data_type = "s"
        cell.number_format = _risk_column_format(section, field, value)
        cell.font = font
        cell.alignment = alignment

    def write_table(worksheet: Any, title_row: int, title: str,
                    columns: tuple[str, ...], records: list[tuple],
                    section: str, start_column: int = 1) -> tuple[int, int]:
        put(worksheet, title_row, start_column, title, section, "section")
        header_row = title_row + 1
        if header_row + len(records) > 1_048_576:
            raise ValueError(f"{worksheet.title}.{title} exceeds the Excel row limit.")
        for column_number, column in enumerate(columns, start=start_column):
            put(worksheet, header_row, column_number, column, section, column)
        for row_number, record in enumerate(records, start=header_row + 1):
            for column_number, (column, value) in enumerate(zip(columns, record), start=start_column):
                put(worksheet, row_number, column_number, value, section, column)
        return header_row, header_row + len(records)

    try:
        for _, period in overview.sort_values("rebalance_no").iterrows():
            number = int(period["rebalance_no"])
            parts = {}
            for name in expected[1:]:
                part = tables[name].loc[tables[name]["rebalance_no"] == number]
                for key in identity_columns:
                    if not part[key].eq(period[key]).all():
                        raise ValueError(f"{name}, rebalance {number}: inconsistent {key}.")
                parts[name] = part
            covariance = parts["Covariance"]
            if len(covariance) != len(ASSET_CLASSES) or (
                set(covariance["asset_class"]) != set(ASSET_CLASSES)
            ):
                raise ValueError(f"Rebalance {number}: covariance must contain four unique classes.")
            covariance = covariance.set_index("asset_class").loc[list(ASSET_CLASSES)]
            if len(parts["ClassReturns"]) != period["n_returns"]:
                raise ValueError(f"Rebalance {number}: return row count differs from n_returns.")
            codes = tuple(parts["Allocations"]["instrument_code"])
            instrument_covariance = parts["InstrumentCovariance"].pivot(
                index="instrument_code", columns="other_instrument_code", values="covariance",
            )
            instrument_returns = parts["InstrumentReturns"].pivot(
                index="return_date", columns="instrument_code", values="return",
            ).sort_index()
            if (len(set(codes)) != len(codes)
                    or set(instrument_covariance.index) != set(codes)
                    or set(instrument_covariance.columns) != set(codes)
                    or set(instrument_returns.columns) != set(codes)):
                raise ValueError(f"Rebalance {number}: instrument tables differ from Allocations.")
            instrument_covariance = instrument_covariance.loc[list(codes), list(codes)]
            instrument_returns = instrument_returns.loc[:, list(codes)]
            if (len(instrument_returns) != period["n_returns"]
                    or not pd.DatetimeIndex(instrument_returns.index).equals(
                        pd.DatetimeIndex(parts["ClassReturns"]["return_date"].sort_values()))
                    or not np.isfinite(instrument_covariance.to_numpy(dtype=float)).all()
                    or not np.isfinite(instrument_returns.to_numpy(dtype=float)).all()):
                raise ValueError(f"Rebalance {number}: instrument observations are incomplete or misaligned.")

            worksheet = workbook.create_sheet(f"Rebalance_{number:03d}")
            worksheet.sheet_view.showGridLines = True
            worksheet.freeze_panes = None
            put(worksheet, 1, 1, f"Rebalance {number:03d}", "Overview", "title")
            for fields, column in ((left_fields, 1), (right_fields, 4)):
                for row_number, field in enumerate(fields, start=3):
                    put(worksheet, row_number, column, field, "Overview", "field")
                    put(worksheet, row_number, column + 1, period[field], "Overview", field)
            put(worksheet, 3, 7, "data_version", "Overview", "field")
            put(worksheet, 3, 8, period["data_version"], "Overview", "data_version")

            covariance_records = [
                (asset, *(covariance.loc[asset, other] for other in ASSET_CLASSES))
                for asset in ASSET_CLASSES
            ]
            write_table(
                worksheet, 14, "AnnualCovariance", ("asset_class", *ASSET_CLASSES),
                covariance_records, "Covariance",
            )
            put(worksheet, 14, 7, "AnnualVolatility", "Overview", "section")
            put(worksheet, 15, 7, "asset_class", "Overview", "field")
            put(worksheet, 15, 8, "annual_volatility", "Overview", "field")
            for row_number, asset in enumerate(ASSET_CLASSES, start=16):
                put(worksheet, row_number, 7, asset, "Overview", "asset_class")
                put(worksheet, row_number, 8, period[f"{asset}_volatility"],
                    "Overview", f"{asset}_volatility")

            allocation_columns = ("asset_class", "instrument_code", "within_class_weight")
            allocation_records = list(
                parts["Allocations"].loc[:, list(allocation_columns)].itertuples(index=False, name=None)
            )
            _, allocation_end = write_table(
                worksheet, 22, "Allocations", allocation_columns,
                allocation_records, "Allocations",
            )
            return_columns = ("return_date", *ASSET_CLASSES)
            return_records = list(
                parts["ClassReturns"].sort_values("return_date").loc[:, list(return_columns)]
                .itertuples(index=False, name=None)
            )
            return_header, return_end = write_table(
                worksheet, allocation_end + 2, "ClassReturns", return_columns,
                return_records, "ClassReturns",
            )
            accounted_for = set(left_fields + right_fields + volatility_fields + ("data_version",))
            diagnostic_fields = [field for field in overview.columns if field not in accounted_for]
            if diagnostic_fields:
                # Preserve Overview fields as explicit field/value pairs. Values
                # keep their field-specific number formats, unlike a generic table.
                diagnostic_start = 22
                put(worksheet, diagnostic_start, 7, "Diagnostics", "Overview", "section")
                put(worksheet, diagnostic_start + 1, 7, "field", "Overview", "field")
                put(worksheet, diagnostic_start + 1, 8, "value", "Overview", "field")
                for row_number, field in enumerate(diagnostic_fields, start=diagnostic_start + 2):
                    put(worksheet, row_number, 7, field, "Overview", "field")
                    put(worksheet, row_number, 8, period[field], "Overview", field)

            write_table(
                worksheet, 14, "InstrumentAnnualCovariance", ("instrument_code", *codes),
                [(code, *instrument_covariance.loc[code].tolist()) for code in codes],
                "InstrumentCovariance", start_column=10,
            )
            write_table(
                worksheet, return_header - 1, "InstrumentReturns", ("return_date", *codes),
                [(day, *values) for day, values in zip(instrument_returns.index, instrument_returns.to_numpy())],
                "InstrumentReturns", start_column=10,
            )

            # These are descriptions of the existing instrument covariance;
            # no matrix repair or additional portfolio constraints are applied.
            matrix = instrument_covariance.to_numpy(dtype=float)
            variances = np.diag(matrix)
            if np.any(variances < 0):
                raise ValueError(f"Rebalance {number}: instrument variance cannot be negative.")
            instrument_volatility = np.sqrt(variances)
            eigenvalues = np.linalg.eigvalsh(matrix)
            rank = int(np.linalg.matrix_rank(matrix))
            condition = float(np.linalg.cond(matrix)) if rank == len(codes) else float("inf")
            statistics_column = len(codes) + 12
            _, volatility_end = write_table(
                worksheet, 14, "InstrumentAnnualVolatility", ("instrument_code", "annual_volatility"),
                list(zip(codes, instrument_volatility)), "InstrumentVolatility",
                start_column=statistics_column,
            )
            diagnostic_start = volatility_end + 2
            put(worksheet, diagnostic_start, statistics_column,
                "InstrumentDiagnostics", "InstrumentDiagnostics", "section")
            put(worksheet, diagnostic_start + 1, statistics_column, "field", "InstrumentDiagnostics", "field")
            put(worksheet, diagnostic_start + 1, statistics_column + 1, "value", "InstrumentDiagnostics", "field")
            for row_number, (field, value) in enumerate((
                ("covariance_rank", rank), ("condition_number", condition),
                ("minimum_eigenvalue", float(eigenvalues[0])),
            ), start=diagnostic_start + 2):
                put(worksheet, row_number, statistics_column, field, "InstrumentDiagnostics", "field")
                put(worksheet, row_number, statistics_column + 1, value, "InstrumentDiagnostics", field)

            # A single regular filter belongs to the daily return table; other
            # sections are independent tables and are deliberately outside it.
            worksheet.auto_filter.ref = f"A{return_header}:E{return_end}"
            for column_number, width in enumerate((28, 22, 24, 25, 22, 3, 26, 22), start=1):
                worksheet.column_dimensions[get_column_letter(column_number)].width = width
            worksheet.column_dimensions["I"].width = 3
            for column_number in range(10, len(codes) + 11):
                worksheet.column_dimensions[get_column_letter(column_number)].width = 22
            worksheet.column_dimensions[get_column_letter(statistics_column - 1)].width = 3
            worksheet.column_dimensions[get_column_letter(statistics_column)].width = 30
            worksheet.column_dimensions[get_column_letter(statistics_column + 1)].width = 22
            worksheet.sheet_properties.pageSetUpPr.fitToPage = False
        workbook.save(destination)
    finally:
        workbook.close()


def save_risk_workbooks(
    tables: dict[str, pd.DataFrame],
    output_root: Union[str, Path],
    lookback_period: str,
    rebalance_frequency: str,
) -> list[Path]:
    """Save one workbook per original rebalance number, for a single scheme.

    Quarterly/monthly results are saved as
    Lookback_Quarterly_Rebalance_Monthly/001.xlsx, 002.xlsx, etc. The number
    identifies the rebalance, never the run. A date-filtered batch retains its
    original numbers. A successful rerun replaces the scheme's numeric XLSX
    files with the current batch, while preserving unrelated files.

    Every workbook is staged before the old directory is replaced. If publishing
    fails, restore the previous directory; leave its backup intact if restoration
    also fails. No original prices or schedule files are modified.
    """
    period_names = {"quarterly": "Quarterly", "semiannual": "Semiannual", "annual": "Annual"}
    frequency_names = {"monthly": "Monthly", "quarterly": "Quarterly", "semiannual": "Semiannual"}
    try:
        folder_name = (
            f"Lookback_{period_names[lookback_period]}_"
            f"Rebalance_{frequency_names[rebalance_frequency]}"
        )
    except KeyError as exc:
        raise ValueError("Invalid lookback period or rebalance frequency.") from exc
    overview = _validate_output_tables(tables)
    if not (
        overview["lookback_period"].eq(lookback_period).all()
        and overview["lookback_days"].eq(LOOKBACK_WINDOWS[lookback_period]).all()
        and overview["rebalance_frequency"].eq(rebalance_frequency).all()
    ):
        raise ValueError("Output tables do not match the selected lookback and rebalance frequency.")
    numbers = sorted(int(number) for number in overview["rebalance_no"])
    metadata = dict(overview.attrs.get("run_metadata", {}))
    root = Path(output_root)
    folder = root / folder_name
    if folder.is_symlink() or (folder.exists() and not folder.is_dir()):
        raise ValueError("The scheme output path must be a regular directory.")
    root.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=".risk-export-", dir=root))
    prepared, previous = staging / "new", staging / "previous"
    prepared.mkdir()
    keep_backup = False
    try:
        for number in numbers:
            period_tables = {
                name: frame.loc[frame["rebalance_no"] == number].copy()
                for name, frame in tables.items()
            }
            period_tables["Overview"].attrs["run_metadata"] = {
                **metadata, "rebalance_no": number,
                "signal_date": period_tables["Overview"].iloc[0]["signal_date"].strftime("%Y-%m-%d"),
                "output_layout": "one_workbook_per_rebalance",
            }
            _write_risk_workbook(period_tables, prepared / f"{number:03d}.xlsx")

        if folder.exists():
            for path in folder.iterdir():
                # Only numbered XLSX files belong to this export batch.
                if path.suffix.lower() == ".xlsx" and re.fullmatch(r"[0-9]+", path.stem):
                    if not path.is_file() or path.is_symlink():
                        raise ValueError(f"A numbered output path is not a regular file: {path}")
                    continue
                target = prepared / path.name
                if path.is_symlink():
                    target.symlink_to(os.readlink(path), target_is_directory=path.is_dir())
                elif path.is_dir():
                    shutil.copytree(path, target, symlinks=True)
                else:
                    shutil.copy2(path, target)
        try:
            if folder.exists():
                os.replace(folder, previous)
            os.replace(prepared, folder)
        except BaseException:
            if previous.exists() and not folder.exists():
                try:
                    os.replace(previous, folder)
                except BaseException as exc:
                    keep_backup = True
                    raise OSError(
                        f"Could not publish or restore outputs. Previous results are preserved at {previous}."
                    ) from exc
            raise
        return [folder / f"{number:03d}.xlsx" for number in numbers]
    finally:
        if not keep_backup and not previous.exists():
            shutil.rmtree(staging, ignore_errors=True)
        elif not keep_backup and folder.exists():
            shutil.rmtree(staging, ignore_errors=True)


def _annualization_days(value: Any) -> int:
    if isinstance(value, bool) or not isinstance(value, Integral) or value < 1:
        raise ValueError("annualization_days must be a positive integer.")
    return int(value)


def run_risk_history(
    prices: pd.DataFrame, schedule: list[dict[str, Any]],
    lookback_period: Optional[str] = None,
    rebalance_frequency: Optional[str] = None,
    start_signal_date: Any = START_SIGNAL_DATE, end_signal_date: Any = END_SIGNAL_DATE,
    annualization_days: Optional[int] = None,
) -> list[dict[str, Any]]:
    """完整运行一个方案，返回逐期结果；任何一期失败都不跳过或保存。"""
    period = LOOKBACK_PERIOD if lookback_period is None else lookback_period
    frequency = REBALANCE_FREQUENCY if rebalance_frequency is None else rebalance_frequency
    if period not in LOOKBACK_WINDOWS or frequency not in INPUT_FREQUENCIES.values():
        raise ValueError("Invalid lookback period or rebalance frequency.")
    factor = _annualization_days(TRADING_DAYS_PER_YEAR if annualization_days is None else annualization_days)
    start = None if start_signal_date is None else _as_date(start_signal_date, "start_signal_date")
    end = None if end_signal_date is None else _as_date(end_signal_date, "end_signal_date")
    if start is not None and end is not None and start > end:
        raise ValueError("start_signal_date must not be after end_signal_date.")
    selected = sorted([
        row for row in schedule if row["lookback"] == LOOKBACK_WINDOWS[period]
        and row["frequency"] == frequency
        and (start is None or row["signal_date"] >= start)
        and (end is None or row["signal_date"] <= end)
    ], key=lambda row: row["signal_date"])
    if not selected:
        raise ValueError("No schedule rows match the selected settings and dates.")
    dates = [row["signal_date"] for row in selected]
    numbers = [row["rebalance_no"] for row in selected]
    if len(set(dates)) != len(dates) or numbers != sorted(set(numbers)):
        raise ValueError("Selected schedule dates and rebalance numbers must be unique and increasing.")
    results = []
    for row in selected:
        try:
            result = estimate_risk(prices, row["signal_date"], row["lookback"], row["allocations"], factor)
            execution = row["execution_date"]
            if execution is not None:
                position = prices.index.get_loc(row["signal_date"])
                if position + 1 >= len(prices) or prices.index[position + 1] != execution:
                    raise ValueError("execution_date must be the next observed trading date.")
        except (ValueError, TypeError) as exc:
            raise ValueError(f"Rebalance {row['rebalance_no']} failed: {exc}") from exc
        result.update({
            "rebalance_no": row["rebalance_no"], "lookback_period": period,
            "rebalance_frequency": frequency, "execution_date": execution,
            "risk_status": "ok",
            "execution_status": "scheduled" if execution is not None else "no_execution_date",
        })
        results.append(result)
    return results


def build_output_tables(
    results: list[dict[str, Any]], metadata: Optional[Mapping[str, Any]] = None,
) -> dict[str, pd.DataFrame]:
    """保留原四张内存表，另加品种协方差和未加权收益表，供按期导出。"""
    if not results:
        raise ValueError("Cannot build output tables without period results.")
    metadata = dict(metadata or {})
    scheme = {(r["lookback_period"], r["rebalance_frequency"]) for r in results}
    if len(scheme) != 1:
        raise ValueError("One export batch must contain exactly one parameter scheme.")
    overview, covariance, allocations, returns = [], [], [], []
    instrument_covariance, instrument_returns = [], []
    keys = ("rebalance_no", "signal_date", "lookback_period", "lookback_days", "rebalance_frequency")
    seen = set()
    for result in results:
        identity = {
            "rebalance_no": result["rebalance_no"], "signal_date": result["signal_date"],
            "lookback_period": result["lookback_period"], "lookback_days": result["lookback"],
            "rebalance_frequency": result["rebalance_frequency"],
        }
        if identity["rebalance_no"] in seen:
            raise ValueError("Duplicate rebalance_no in output results.")
        seen.add(identity["rebalance_no"])
        diagnostics = result["diagnostics"]
        overview.append({
            **identity, "execution_date": result["execution_date"],
            **{key: result[key] for key in ("price_start", "price_end", "return_start", "return_end", "n_prices", "n_returns")},
            **{f"{asset}_volatility": float(result["annual_volatility"][asset]) for asset in ASSET_CLASSES},
            "risk_status": result["risk_status"], "execution_status": result["execution_status"],
            "covariance_rank": diagnostics["rank"], "condition_number": diagnostics["condition_number"],
            "minimum_eigenvalue": float(diagnostics["eigenvalues"][0]),
            "annualization_days": diagnostics["annualization_factor"],
            "covariance_ddof": 1, "return_method": "simple",
            "run_date": pd.Timestamp(metadata["run_date"]) if metadata.get("run_date") else None,
            "data_version": metadata.get("data_version", "unspecified"),
        })
        for asset in ASSET_CLASSES:
            covariance.append({
                **identity, "asset_class": asset,
                **{other: float(result["covariance"].loc[asset, other]) for other in ASSET_CLASSES},
            })
            for code, weight in result["allocations"][asset].items():
                allocations.append({**identity, "asset_class": asset, "instrument_code": code, "within_class_weight": weight})
        for day, values in result["class_returns"].iterrows():
            returns.append({**identity, "return_date": day, **{asset: float(values[asset]) for asset in ASSET_CLASSES}})
        for code, values in result["instrument_covariance"].iterrows():
            for other, value in values.items():
                instrument_covariance.append({**identity, "instrument_code": code,
                                              "other_instrument_code": other, "covariance": float(value)})
        for day, values in result["instrument_returns"].iterrows():
            for code, value in values.items():
                instrument_returns.append({**identity, "return_date": day,
                                           "instrument_code": code, "return": float(value)})
    tables = {
        "Overview": pd.DataFrame(overview),
        "Covariance": pd.DataFrame(covariance, columns=[*keys, "asset_class", *ASSET_CLASSES]),
        "Allocations": pd.DataFrame(allocations, columns=[*keys, "asset_class", "instrument_code", "within_class_weight"]),
        "ClassReturns": pd.DataFrame(returns, columns=[*keys, "return_date", *ASSET_CLASSES]),
        "InstrumentCovariance": pd.DataFrame(instrument_covariance, columns=[
            *keys, "instrument_code", "other_instrument_code", "covariance"]),
        "InstrumentReturns": pd.DataFrame(instrument_returns, columns=[
            *keys, "return_date", "instrument_code", "return"]),
    }
    tables["Overview"].attrs["run_metadata"] = metadata
    return tables


def _file_digest(path: PathLike) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Estimate full-history risk and save one XLSX workbook per rebalance.")
    parser.add_argument("--prices", type=Path, default=DEFAULT_PRICE_PATH)
    parser.add_argument("--schedule", type=Path, default=DEFAULT_SCHEDULE_PATH)
    parser.add_argument("--lookback-period", choices=tuple(LOOKBACK_WINDOWS), default=LOOKBACK_PERIOD)
    parser.add_argument("--frequency", choices=tuple(INPUT_FREQUENCIES.values()), default=REBALANCE_FREQUENCY)
    parser.add_argument("--start-signal-date", default=START_SIGNAL_DATE)
    parser.add_argument("--end-signal-date", default=END_SIGNAL_DATE)
    parser.add_argument("--annualization-days", type=int, default=TRADING_DAYS_PER_YEAR)
    parser.add_argument("--output-root", type=Path, default=OUTPUT_ROOT)
    args = parser.parse_args(argv)
    try:
        price_hash, schedule_hash = _file_digest(args.prices), _file_digest(args.schedule)
        prices, schedule = load_prices(args.prices), load_schedule(args.schedule)
        if price_hash != _file_digest(args.prices) or schedule_hash != _file_digest(args.schedule):
            raise ValueError("An input file changed while it was being loaded; run again after saving inputs.")
        results = run_risk_history(prices, schedule, args.lookback_period, args.frequency,
                                   args.start_signal_date, args.end_signal_date, args.annualization_days)
        metadata = {
            "schema_version": "3.0", "output_layout": "one_workbook_per_rebalance",
            "worksheet_layout": "one_rebalance_per_sheet",
            "signal_date_definition": "Last trading date whose close is included in the estimation window", "run_date": datetime.now(timezone(timedelta(hours=8))).date().isoformat(),
            "lookback_period": args.lookback_period, "lookback_days": LOOKBACK_WINDOWS[args.lookback_period],
            "rebalance_frequency": args.frequency, "annualization_days": args.annualization_days,
            "covariance_ddof": 1, "return_method": "simple", "asset_order": list(ASSET_CLASSES),
            "price_source": str(args.prices.resolve()), "schedule_source": str(args.schedule.resolve()),
            "price_sha256": price_hash, "schedule_sha256": schedule_hash,
            "script_sha256": _file_digest(__file__),
            "data_version": hashlib.sha256(f"{price_hash}|{schedule_hash}".encode()).hexdigest()[:16],
            "start_signal_date": args.start_signal_date, "end_signal_date": args.end_signal_date,
        }
        tables = build_output_tables(results, metadata)
        outputs = save_risk_workbooks(tables, args.output_root, args.lookback_period, args.frequency)
    except (ValueError, OSError, ImportError, TypeError) as exc:
        parser.exit(1, f"Risk estimation failed: {exc}\n")
    print(f"Lookback period: {args.lookback_period} ({LOOKBACK_WINDOWS[args.lookback_period]} daily returns)")
    print(f"Rebalance frequency: {args.frequency}")
    print(f"Periods processed: {len(results)}")
    print(f"Signal dates: {results[0]['signal_date']:%Y-%m-%d} to {results[-1]['signal_date']:%Y-%m-%d}")
    print(f"Signals without an execution date: {sum(r['execution_date'] is None for r in results)}")
    print(f"Workbooks saved: {len(outputs)} (one per rebalance)")
    print(f"Output directory: {outputs[0].parent.resolve()}")
    print(f"First file: {outputs[0].name}; last file: {outputs[-1].name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
