"""读取已验收的逐期优化 Excel，导出回测及目标波动率比较 Excel。

Python 3.9+；只用 numpy/pandas/openpyxl。导入不读写，不改仓位、不生成 JSON。
未定义统计量保存为文字 N/A；图表将其视为缺失，不伪造为零。
"""
from __future__ import annotations

import hashlib
import math
import os
import re
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

GROUP_ORDER = ("equity", "bonds", "gold", "other_commodities")
SUCCESS_STATUSES = {"target_achieved", "cap_limited"}
ABS_TOL = 1e-12
REL_TOL = 1e-10


def _number(value: Any, name: str, positive: bool = False) -> float:
    if isinstance(value, (bool, np.bool_)) or not isinstance(value, (int, float, np.number)):
        raise ValueError(f"{name} must be numeric")
    value = float(value)
    if not math.isfinite(value) or (positive and value <= 0):
        raise ValueError(f"{name} must be finite" + (" and positive" if positive else ""))
    return value


def _date(value: Any, name: str, nullable: bool = False):
    if value is None or value == "":
        if nullable:
            return None
        raise ValueError(f"Missing {name}")
    if isinstance(value, (bool, int, float, np.number)):
        raise ValueError(f"{name} must be a date")
    if isinstance(value, str) and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        raise ValueError(f"{name} must use YYYY-MM-DD")
    try:
        result = pd.Timestamp(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid {name}") from exc
    if pd.isna(result) or result.tzinfo is not None or result != result.normalize():
        raise ValueError(f"Invalid {name}")
    return result


def _same_number(actual: Any, expected: Any, name: str, atol: float = ABS_TOL):
    actual, expected = _number(actual, name), _number(expected, name)
    if abs(actual - expected) > atol + REL_TOL * abs(expected):
        raise ValueError(f"Inconsistent {name}: {actual!r} != {expected!r}")


def _digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _rows(book, sheet: str, required: tuple) -> list[dict]:
    if sheet not in book.sheetnames:
        raise ValueError(f"Missing worksheet {sheet}")
    rows = []
    for row in book[sheet].iter_rows():
        if any(c.data_type == "f" for c in row):
            raise ValueError(f"Formula in numeric input worksheet {sheet}")
        rows.append(tuple(c.value for c in row))
    if not rows:
        raise ValueError(f"Empty worksheet {sheet}")
    headers = list(rows[0])
    while headers and headers[-1] is None:
        headers.pop()
    if (not headers or any(not isinstance(h, str) or not h or h != h.strip() for h in headers)
            or len(set(headers)) != len(headers) or not set(required).issubset(headers)):
        raise ValueError(f"Invalid headers in {sheet}")
    output = []
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        if any(v is not None for v in row[len(headers):]):
            raise ValueError(f"Unlabelled data in {sheet}")
        output.append({h: row[j] if j < len(row) else None for j, h in enumerate(headers)})
    if not output:
        raise ValueError(f"Empty data in {sheet}")
    return output


def _fields(rows: list[dict], unique: bool) -> dict[str, list]:
    result = {}
    for row in rows:
        key = row["field"]
        if not isinstance(key, str) or not key or key != key.strip():
            raise ValueError("Invalid field name")
        if unique and key in result:
            raise ValueError(f"Duplicate field {key}")
        result.setdefault(key, []).append(row["value"])
    return result


def _input(values: dict, name: str, kind: str = "number", nullable: bool = False):
    if name not in values:
        raise ValueError(f"Missing Inputs field {name}")
    normalized = []
    for raw in values[name]:
        if kind == "date":
            normalized.append(_date(raw, name, nullable))
        elif kind == "number":
            normalized.append(_number(raw, name))
        else:
            if not isinstance(raw, str) or not raw:
                raise ValueError(f"Invalid Inputs field {name}")
            normalized.append(raw)
    first = normalized[0]
    for other in normalized[1:]:
        if kind == "number":
            _same_number(other, first, f"duplicate Inputs {name}")
        elif other != first:
            raise ValueError(f"Conflicting duplicate Inputs field {name}")
    return first


def read_optimization(path: Path, expected: dict, target: float, prices=None, *, leverage_cap: float) -> dict:
    """读取一个原始调仓结果；失败、缺字段、冲突、非有限值均拒绝。"""
    path = Path(path)
    target = _number(target, "target", True)
    leverage_cap = _number(leverage_cap, "bond_leverage_cap", True)
    initial_hash = _digest(path)
    book = load_workbook(path, read_only=True, data_only=False)
    try:
        overview = {k: v[0] for k, v in _fields(_rows(book, "Overview", ("field", "value")), True).items()}
        required = {"status", "rebalance_no", "signal_date", "execution_date", "execution_status",
                    "target_volatility", "actual_volatility", "non_bond_exposure", "gross_exposure",
                    "objective", "volatility_gap", "cap_multiplier", "bond_leverage_cap", "bond_margin_floor",
                    "bond_margin_rate", "bond_leverage", "bond_exposure", "bond_capital_weight",
                    "capital_used_at_floor", "capital_usage"}
        if required - overview.keys():
            raise ValueError(f"Missing Overview fields: {sorted(required - overview.keys())}")
        status = overview["status"]
        if status not in SUCCESS_STATUSES:
            raise ValueError(f"Optimization {path.name} is not accepted: {status}")
        number = _number(overview["rebalance_no"], "rebalance_no", True)
        expected_number = _number(expected["rebalance_no"], "expected rebalance_no", True)
        if (not number.is_integer() or number != expected_number or not re.fullmatch(r"[0-9]+", path.stem)
                or int(path.stem) != int(number)):
            raise ValueError("Optimization filename/sequence/schedule mismatch")
        signal = _date(overview["signal_date"], "signal_date")
        execution = _date(overview["execution_date"], "execution_date", True)
        if signal != _date(expected["signal_date"], "expected signal_date"):
            raise ValueError("Optimization signal date differs from schedule")
        if execution != _date(expected["execution_date"], "expected execution_date", True):
            raise ValueError("Optimization execution date differs from schedule")
        execution_status = "no_execution_date" if execution is None else "scheduled"
        if overview["execution_status"] != execution_status or (execution is not None and execution <= signal):
            raise ValueError("Invalid optimization execution date/status")
        _same_number(overview["target_volatility"], target, "target_volatility")
        sigma = _number(overview["actual_volatility"], "actual_volatility", True)
        objective = _number(overview["objective"], "objective")
        _same_number(objective, (sigma - target) ** 2, "objective")
        _same_number(overview["volatility_gap"], sigma - target, "volatility_gap")
        non_bond = _number(overview["non_bond_exposure"], "non_bond_exposure", True)
        cap = _number(overview["capital_used_at_floor"], "capital_used_at_floor", True)
        floor = _number(overview["bond_margin_floor"], "bond_margin_floor", True)
        margin = _number(overview["bond_margin_rate"], "bond_margin_rate", True)
        bond = _number(overview["bond_exposure"], "bond_exposure", True)
        _same_number(overview["bond_leverage_cap"], leverage_cap, "bond_leverage_cap")
        _same_number(floor, 1.0 / leverage_cap, "bond_margin_floor")
        _same_number(overview["bond_leverage"], 1.0 / margin, "bond_leverage")
        _same_number(overview["bond_capital_weight"], margin * bond, "bond_capital_weight")
        _same_number(cap, non_bond + floor * bond, "capital_used_at_floor")
        _same_number(overview["capital_usage"], non_bond + margin * bond, "capital_usage")
        _same_number(overview["capital_usage"], 1., "capital_budget")
        gross = _number(overview["gross_exposure"], "gross_exposure", True)
        multiplier = _number(overview["cap_multiplier"], "cap_multiplier")
        inputs = _fields(_rows(book, "Inputs", ("field", "value")), False)
        for key, value in (("rebalance_no", number), ("lookback_days", expected["lookback"]),
                           ("target_volatility", target), ("bond_leverage_cap", leverage_cap), ("bond_margin_floor", floor)):
            _same_number(_input(inputs, key), value, key)
        if _input(inputs, "rebalance_frequency", "text") != expected["frequency"]:
            raise ValueError("Optimization frequency differs from schedule")
        if _input(inputs, "signal_date", "date") != signal or _input(inputs, "execution_date", "date", True) != execution:
            raise ValueError("Overview/Inputs dates disagree")
        if _input(inputs, "execution_status", "text") != execution_status:
            raise ValueError("Overview/Inputs execution status disagree")
        if _input(inputs, "group_order", "text") != ", ".join(GROUP_ORDER):
            raise ValueError("Unexpected group order")
        if _input(inputs, "return_method", "text") != "simple":
            raise ValueError("Optimization uses a different return convention")
        _same_number(_input(inputs, "annualization_days"), 252, "annualization_days")
        tolerances = {key: _input(inputs, key) for key in ("budget_tol", "cap_tol", "vol_tol", "residual_tol")}
        if any(v <= 0 for v in tolerances.values()):
            raise ValueError("Optimization tolerances must be positive")
        if cap > 1 + tolerances["cap_tol"] or multiplier < -tolerances["residual_tol"]:
            raise ValueError("Invalid cap exposure or cap multiplier")
        if (floor - margin) * bond > tolerances["cap_tol"]:
            raise ValueError("Bond margin is below the selected floor")
        if abs(multiplier * (cap - 1)) > tolerances["cap_tol"]:
            raise ValueError("Optimization fails complementarity")
        if status == "target_achieved" and abs(sigma - target) > tolerances["vol_tol"]:
            raise ValueError("target_achieved is inconsistent with actual volatility")
        if status == "cap_limited" and (abs(cap - 1) > tolerances["cap_tol"] or sigma >= target or multiplier <= 0):
            raise ValueError("cap_limited is inconsistent with boundary/volatility/multiplier")
        check_rows = _rows(book, "Checks", ("check", "error", "tolerance", "passed", "detail"))
        check_names = []
        for item in check_rows:
            name = item["check"]
            if not isinstance(name, str) or not name or name in check_names:
                raise ValueError("Invalid/duplicate optimization check")
            check_names.append(name)
            if not isinstance(item["passed"], (bool, np.bool_)) or not bool(item["passed"]):
                raise ValueError(f"Optimization check did not pass: {name}")
            error = _number(item["error"], f"check {name} error")
            check_tol = _number(item["tolerance"], f"check {name} tolerance")
            if check_tol < 0:
                raise ValueError(f"Negative tolerance in check {name}")
            error_checks = {"all_four_risk_shares", "capital_cap_at_floor", "stationarity", "nonnegative_cap_multiplier",
                            "complementarity", "variance_groups_equals_instruments", "group_risk_contributions_sum",
                            "instrument_risk_contributions_sum", "target_achieved", "active_cap_equality",
                            "bond_margin_floor", "bond_leverage_cap", "capital_budget_equality", "group_capital_sum", "instrument_capital_sum"}
            error_checks.update(f"{g}_instrument_contributions" for g in GROUP_ORDER)
            error_checks.update(f"{g}_instrument_capital" for g in GROUP_ORDER)
            if (name in error_checks or name.startswith("solver_")) and (error < 0 or error > check_tol):
                raise ValueError(f"Optimization check error contradicts passed=True: {name}")
            if name in {"positive_group_exposures", "positive_bond_capital", "cap_limited_positive_multiplier"} and error <= 0:
                raise ValueError(f"Positive-value optimization check contradicts passed=True: {name}")
            if name == "cap_limited_below_target" and error >= 0:
                raise ValueError("Below-target optimization check contradicts passed=True")
        required_checks = {"saved_class_returns_match_prices", "saved_Q_matches_class_returns", "saved_Q_matches_A_Sigma_A",
                           "positive_group_exposures", "all_four_risk_shares", "capital_cap_at_floor", "stationarity",
                           "nonnegative_cap_multiplier", "complementarity", "variance_groups_equals_instruments",
                           "group_risk_contributions_sum", "instrument_risk_contributions_sum", "positive_bond_capital",
                           "bond_margin_floor", "bond_leverage_cap", "capital_budget_equality", "group_capital_sum", "instrument_capital_sum",
                           "saved_instrument_returns_match_prices", "saved_Sigma_matches_instrument_returns",
                           "saved_class_returns_match_instruments", "saved_class_volatility_matches_Q", "saved_instrument_volatility_matches_Sigma"}
        required_checks.update(f"{g}_instrument_contributions" for g in GROUP_ORDER)
        required_checks.update(f"{g}_instrument_capital" for g in GROUP_ORDER)
        required_checks.update({"target_achieved"} if status == "target_achieved" else
                               {"active_cap_equality", "cap_limited_below_target", "cap_limited_positive_multiplier"})
        if required_checks - set(check_names):
            raise ValueError(f"Missing optimization checks: {sorted(required_checks - set(check_names))}")
        group_rows = _rows(book, "GroupResults", ("asset_class", "target_exposure", "budget", "risk_share", "budget_error",
                                                   "marginal_risk", "risk_contribution", "variance_contribution", "budget_multiplier", "capital_weight"))
        group_results = {}
        for item in group_rows:
            group = item["asset_class"]
            if group not in GROUP_ORDER or group in group_results:
                raise ValueError("Invalid/duplicate group result")
            for key in ("target_exposure", "budget", "risk_share", "budget_error", "marginal_risk", "risk_contribution", "variance_contribution"):
                item[key] = _number(item[key], f"{group} {key}", key in {"target_exposure", "budget"})
            item["capital_weight"] = _number(item["capital_weight"], f"{group} capital_weight", True)
            _same_number(item["capital_weight"], item["target_exposure"] * (margin if group == "bonds" else 1.), f"{group} capital mapping")
            if group != "other_commodities" or item["budget_multiplier"] is not None:
                _number(item["budget_multiplier"], f"{group} budget multiplier")
            _same_number(item["budget"], _input(inputs, f"budget_{group}"), f"{group} budget")
            if abs(item["risk_share"] - item["budget"]) > tolerances["budget_tol"]:
                raise ValueError(f"Risk budget fails for {group}")
            _same_number(item["budget_error"], item["risk_share"] - item["budget"], f"{group} budget_error")
            _same_number(item["risk_contribution"], item["target_exposure"] * item["marginal_risk"], f"{group} risk contribution")
            _same_number(item["risk_share"], item["risk_contribution"] / sigma, f"{group} risk share")
            _same_number(item["variance_contribution"], item["risk_contribution"] * sigma, f"{group} variance contribution")
            _same_number(_input(inputs, f"h_{group}"), floor if group == "bonds" else 1, f"h_{group}")
            group_results[group] = item
        if set(group_results) != set(GROUP_ORDER) or set(expected["allocations"]) != set(GROUP_ORDER):
            raise ValueError("Exactly four groups are required")
        _same_number(math.fsum(r["budget"] for r in group_results.values()), 1, "budget sum")
        _same_number(math.fsum(r["target_exposure"] for r in group_results.values()), gross, "gross exposure")
        _same_number(math.fsum(r["target_exposure"] for g, r in group_results.items() if g != "bonds"), non_bond, "non-bond exposure")
        _same_number(group_results["bonds"]["target_exposure"], bond, "bond exposure")
        _same_number(math.fsum(r["capital_weight"] for r in group_results.values()), 1., "group capital sum")
        _same_number(math.fsum(r["risk_contribution"] for r in group_results.values()), sigma, "group risk sum")
        instrument_rows = _rows(book, "InstrumentResults", ("instrument_code", "asset_class", "included", "within_class_weight",
                                                             "target_exposure", "marginal_risk", "risk_contribution", "risk_share", "capital_weight"))
        exposures, groups, capital_weights = {}, {}, {}
        rows_by_group = {g: [] for g in GROUP_ORDER}
        for item in instrument_rows:
            code, group = item["instrument_code"], item["asset_class"]
            if (not isinstance(code, str) or not code or code != code.strip() or code in exposures or group not in GROUP_ORDER
                    or not isinstance(item["included"], (bool, np.bool_)) or not bool(item["included"])):
                raise ValueError("Invalid/duplicate/excluded instrument result")
            planned = expected["allocations"][group]
            if code not in planned:
                raise ValueError(f"Unexpected instrument {code}")
            for key in ("within_class_weight", "target_exposure", "marginal_risk", "risk_contribution", "risk_share"):
                item[key] = _number(item[key], f"{code} {key}", key in {"within_class_weight", "target_exposure"})
            item["capital_weight"] = _number(item["capital_weight"], f"{code} capital_weight", True)
            _same_number(item["capital_weight"], item["target_exposure"] * (margin if group == "bonds" else 1.), f"{code} capital mapping")
            capital_weights[code] = item["capital_weight"]
            _same_number(item["within_class_weight"], planned[code], f"{code} scheduled within-class weight")
            _same_number(item["within_class_weight"], 1 / len(planned), f"{code} equal within-class weight")
            _same_number(item["target_exposure"], item["within_class_weight"] * group_results[group]["target_exposure"], f"{code} mapped exposure")
            _same_number(item["risk_contribution"], item["target_exposure"] * item["marginal_risk"], f"{code} risk contribution")
            _same_number(item["risk_share"], item["risk_contribution"] / sigma, f"{code} risk share")
            exposures[code], groups[code] = item["target_exposure"], group
            rows_by_group[group].append(item)
        for group, items in rows_by_group.items():
            if {r["instrument_code"] for r in items} != set(expected["allocations"][group]):
                raise ValueError(f"Missing scheduled instruments in {group}")
            _same_number(math.fsum(r["target_exposure"] for r in items), group_results[group]["target_exposure"], f"{group} instrument exposure sum")
            _same_number(math.fsum(r["risk_contribution"] for r in items), group_results[group]["risk_contribution"], f"{group} instrument risk sum")
            _same_number(math.fsum(r["capital_weight"] for r in items), group_results[group]["capital_weight"], f"{group} instrument capital sum")
        if prices is not None:
            if (not isinstance(prices, pd.DataFrame) or not isinstance(prices.index, pd.DatetimeIndex)
                    or prices.index.has_duplicates or not prices.index.is_monotonic_increasing
                    or prices.columns.has_duplicates or signal not in prices.index):
                raise ValueError("Invalid price table or missing signal date")
            lookback = _number(expected["lookback"], "expected lookback", True)
            if not lookback.is_integer():
                raise ValueError("lookback must be an integer")
            lookback = int(lookback)
            end = prices.index.get_loc(signal)
            codes = list(exposures)
            if end < lookback or any(code not in prices for code in codes):
                raise ValueError("Insufficient price history or missing instrument")
            window = prices.loc[:, codes].iloc[end - lookback:end + 1]
            if any(pd.api.types.is_bool_dtype(window[c]) for c in codes):
                raise ValueError("Boolean prices are invalid")
            values = window.to_numpy(dtype=float)
            if not np.isfinite(values).all() or np.any(values <= 0):
                raise ValueError("Missing, nonfinite or nonpositive prices in optimization window")
            with np.errstate(over="raise", divide="raise", invalid="raise"):
                rebuilt = values[1:] / values[:-1] - 1
            saved = _rows(book, "InstrumentReturns", ("return_date", *codes))
            if len(saved) != lookback or set(saved[0]) != {"return_date", *codes}:
                raise ValueError("Unexpected optimization return window or instrument columns")
            saved_dates = pd.DatetimeIndex([_date(r["return_date"], "saved return date") for r in saved])
            if not saved_dates.equals(window.index[1:]):
                raise ValueError("Optimization return dates disagree with price history")
            saved_values = np.array([[_number(r[c], f"saved return {c}") for c in codes] for r in saved])
            if not np.allclose(rebuilt, saved_values, atol=ABS_TOL, rtol=REL_TOL):
                raise ValueError("Optimization historical returns differ from current price history")
        source_price_sha256 = _input(inputs, "price_sha256", "text")
        if not re.fullmatch(r"[0-9a-f]{64}", source_price_sha256):
            raise ValueError("Invalid source price hash")
    finally:
        book.close()
    if _digest(path) != initial_hash:
        raise ValueError("Optimization input changed while being read")
    return {"rebalance_no": int(number), "signal_date": signal, "execution_date": execution,
            "status": status, "target_volatility": target, "actual_volatility": sigma,
            "exposures": exposures, "groups": groups,
            "capital_weights": capital_weights, "bond_margin_floor": floor, "bond_leverage_cap": leverage_cap,
            "bond_margin_rate": margin, "bond_leverage": 1.0 / margin,
            "bond_capital_weight": margin * bond,
            "group_exposures": {g: group_results[g]["target_exposure"] for g in GROUP_ORDER},
            "source_path": path.resolve(), "source_sha256": initial_hash, "source_price_sha256": source_price_sha256}


# 展示层沿用 ETF 回测的独立 Excel 文件和样式；图片由绘图模块单独输出。
_BLANK = object()  # 未估计的容量和基线收益留空，与未定义统计量 N/A 区分。
GROUP_NAMES = {"equity": "股票", "bonds": "债券", "gold": "黄金", "other_commodities": "其他商品"}
FUNDING_COLUMNS = {
    "bond_leverage_cap": "国债杠杆上限", "bond_margin_floor": "最低保证金率",
    "target_bond_margin_rate": "最近调仓目标保证金率", "bond_notional": "国债名义金额",
    "bond_capital": "国债资金", "bond_capital_weight": "国债资金占比",
    "bond_margin_rate": "当日国债保证金率", "bond_leverage": "当日国债杠杆",
    "minimum_bond_capital": "最低国债资金要求", "capital_buffer": "国债资金余量",
    "capital_shortfall": "国债资金缺口", "capital_used_at_floor": "按下限计算的资金占用",
    "capital_usage": "资金分配合计", "margin_breach": "每日收盘是否越界",
    "bond_capital_depleted": "每日收盘国债资金是否耗尽", "funding_state": "每日收盘资金状态",
    "bond_capital_before_trade": "交易前国债资金", "bond_margin_rate_before_trade": "交易前国债保证金率",
    "bond_leverage_before_trade": "交易前国债杠杆", "capital_shortfall_before_trade": "交易前国债资金缺口",
    "margin_breach_before_trade": "交易前是否越界", "bond_capital_depleted_before_trade": "交易前国债资金是否耗尽",
    "margin_breach_observed": "当日是否观察到越界",
}
FUNDING_PERCENT = {FUNDING_COLUMNS[k] for k in ("bond_margin_floor", "target_bond_margin_rate", "bond_capital_weight",
                   "bond_margin_rate", "bond_margin_rate_before_trade", "capital_used_at_floor", "capital_usage")}
FUNDING_DECIMAL = {FUNDING_COLUMNS[k] for k in ("bond_leverage_cap", "bond_notional", "bond_capital", "bond_leverage",
                   "minimum_bond_capital", "capital_buffer", "capital_shortfall", "bond_capital_before_trade",
                   "bond_leverage_before_trade", "capital_shortfall_before_trade")}


def _funding_values(record):
    labels = {"not_invested": "尚未建仓", "within_limit": "满足下限", "below_floor": "低于下限",
              "bond_capital_depleted": "国债资金耗尽"}
    values = []
    for key in FUNDING_COLUMNS:
        value = record.get(key)
        if isinstance(value, (bool, np.bool_)):
            value = "是" if value else "否"
        elif key == "funding_state":
            value = labels.get(value, value)
        values.append(value)
    return values
ANNUAL_HEADERS = ["年份", "策略收益", "基准收益", "超额收益", "年化波动", "跟踪误差", "Sharpe", "Sortino",
                  "信息比率", "策略最大回撤", "基准最大回撤", "超额最大回撤", "Calmar", "年化换手率（单边）"]


def _scalar(value: Any):
    if value is _BLANK:
        return None
    if value is None or value is pd.NA or value is pd.NaT:
        return "N/A"
    if isinstance(value, np.generic):
        value = value.item()
    if isinstance(value, pd.Timestamp):
        if value.tzinfo is not None:
            raise ValueError("Excel dates must be timezone-naive")
        return value.to_pydatetime()
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value):
            return "N/A"
        if not math.isfinite(value):
            raise ValueError("Cannot export infinite numeric values")
        return value
    if isinstance(value, (str, bool, int, datetime, date)):
        return value
    if isinstance(value, (tuple, list)):
        return "、".join(str(v) for v in value)
    raise TypeError(f"Unsupported Excel value: {type(value).__name__}")


def _missing(value):
    return value is None or value is pd.NA or value is pd.NaT or (isinstance(value, (float, np.floating)) and math.isnan(float(value)))


def style_worksheet(sheet, percent_headers=None, decimal_headers=None, integer_headers=None, max_width=28):
    """复用 ETF：Arial、4472C4 蓝表头、灰边框、A2 冻结、筛选、隐藏网格。"""
    percent_headers, decimal_headers, integer_headers = set(percent_headers or ()), set(decimal_headers or ()), set(integer_headers or ())
    thin_gray = Side(style="thin", color="D9D9D9")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.font = Font(name="Arial", size=11, color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    sheet.row_dimensions[1].height = 24
    headers = {cell.value: cell.column for cell in sheet[1]}
    for names, number_format in ((percent_headers, "0.00%"), (decimal_headers, "0.000000"), (integer_headers, "#,##0")):
        for header in names:
            if header in headers:
                for row in range(2, sheet.max_row + 1):
                    sheet.cell(row, headers[header]).number_format = number_format
    for column_number in range(1, sheet.max_column + 1):
        values = [sheet.cell(row, column_number).value for row in range(1, min(sheet.max_row, 300) + 1)]
        width = min(max((len(str(v)) for v in values if v is not None), default=10) + 2, max_width)
        sheet.column_dimensions[get_column_letter(column_number)].width = max(width, 11)
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = Alignment(horizontal="center" if isinstance(cell.value, str) else "right", vertical="center")
            if isinstance(cell.value, (datetime, date)):
                cell.number_format = "yyyy-mm-dd"


def _table(book, name, headers, records, *, percent=(), decimal=(), integer=(), max_width=28):
    headers = list(headers)
    if not headers or len(set(headers)) != len(headers) or any(not isinstance(h, str) or not h for h in headers):
        raise ValueError(f"Invalid headers for {name}")
    if len(headers) > 16384:
        raise ValueError(f"{name} exceeds Excel column limit")
    ws = book.create_sheet(name)
    ws.append(headers)
    count = 0
    for record in records:
        values = list(record)
        if len(values) != len(headers):
            raise ValueError(f"Row length mismatch in {name}")
        count += 1
        if count > 1048575:
            raise ValueError(f"{name} exceeds Excel row limit")
        ws.append([_scalar(v) for v in values])
    for row in ws:
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"  # 包括以 = 开头的路径/说明；禁止作为 Excel 公式执行。
    style_worksheet(ws, percent, decimal, integer, max_width)
    return ws


def _new_book():
    book = Workbook()
    book.remove(book.active)
    return book


def _validate_destination(path: Path, overwrite: bool):
    if not isinstance(overwrite, bool):
        raise TypeError("overwrite must be bool")
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Output must use .xlsx")
    for component in (path, *path.absolute().parents):
        if component.is_symlink():
            raise ValueError(f"Refusing symbolic-link output path: {component}")
    if path.exists() and (not overwrite or not path.is_file()):
        raise FileExistsError(f"Output already exists or is not a file: {path}")


def _save(book, path, overwrite):
    path, temporary = Path(path), None
    try:
        _validate_destination(path, overwrite)
        path.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix=".backtest-", suffix=".xlsx", dir=path.parent)
        os.close(fd)
        book.save(temporary)
        _validate_destination(path, overwrite)
        if overwrite:
            os.replace(temporary, path)
        else:
            os.link(temporary, path)
    finally:
        book.close()
        if temporary is not None:
            Path(temporary).unlink(missing_ok=True)


def report_paths(output_dir, mode="next_day_close") -> list[Path]:
    if not isinstance(mode, str) or not re.fullmatch(r"[a-z][a-z0-9_]*", mode):
        raise ValueError("Invalid execution mode filename")
    root = Path(output_dir)
    return [root / f"{mode}_{suffix}.xlsx" for suffix in
            ("annual_metrics", "backtest_metrics", "holdings", "time_series", "account_details")]


def _frame(result, key, required=()):
    frame = result[key]
    if not isinstance(frame, pd.DataFrame) or frame.columns.has_duplicates:
        raise ValueError(f"{key} must be a DataFrame with unique columns")
    if len(frame) and not set(required).issubset(frame.columns):
        raise ValueError(f"Missing required {key} columns: {sorted(set(required) - set(frame.columns))}")
    return frame


def _nav_column(frame):
    for column in ("normalized_nav", "net_nav", "nav"):
        if column in frame:
            return column
    raise ValueError("Daily NAV is missing")


def _date_text(value):
    return None if _missing(value) else _date(value, "report date").strftime("%Y-%m-%d")


def _drawdown_range(summary):
    start, end = summary.get("max_drawdown_start"), summary.get("max_drawdown_end")
    if _missing(start) or _missing(end):
        return None
    return f"{_date(start, 'drawdown start'):%Y年%m月}-{_date(end, 'drawdown end'):%Y年%m月}"


def _value(record, *keys):
    for key in keys:
        if key in record:
            return record[key]
    return None


def _performance_rows(result):
    s = result["summary"]
    signals = _frame(result, "signals")
    attempted = int((signals["backtest_action"] == "executed").sum()) if "backtest_action" in signals else 0
    success_rate = s.get("rebalance_count", 0) / attempted if attempted else None
    return [
        ["基本信息", "回测区间", f"{_date_text(s['start_date'])}至{_date_text(s['end_date'])}"],
        ["基本信息", "交易日数量", s.get("observations")],
        ["基本信息", "初始净值", 1.0],
        ["基本信息", "期末净值", s["final_nav"] / s["initial_nav"]],
        ["基本信息", "累计收益率", s.get("total_return")],
        ["收益表现", "策略年化收益率", s.get("annualized_return")],
        ["收益表现", "基准年化收益率", s.get("benchmark_annualized_return")],
        ["收益表现", "年化超额收益", s.get("annualized_excess_return")],
        ["收益表现", "基准年化波动率", s.get("benchmark_annualized_volatility")],
        ["收益表现", "基准最大回撤", s.get("benchmark_max_drawdown")],
        ["绝对风险收益", "策略年化波动率", s.get("annualized_volatility")],
        ["绝对风险收益", "Sharpe", s.get("sharpe_ratio")],
        ["绝对风险收益", "Sortino", s.get("sortino_ratio")],
        ["绝对风险收益", "策略最大回撤", s.get("max_drawdown")],
        ["绝对风险收益", "策略最大回撤区间", _drawdown_range(s)],
        ["绝对风险收益", "Calmar", s.get("calmar_ratio")],
        ["相对基准表现", "年化跟踪误差", s.get("tracking_error")],
        ["相对基准表现", "IR", s.get("information_ratio")],
        ["相对基准表现", "超额最大回撤", s.get("excess_max_drawdown")],
        ["相对基准表现", "超额最大回撤区间", _drawdown_range({"max_drawdown_start": s.get("excess_max_drawdown_start"), "max_drawdown_end": s.get("excess_max_drawdown_end")})],
        ["交易与组合", "年化换手率（单边）", s.get("annualized_one_way_turnover")],
        ["交易与组合", "平均每日换仓比率（单边）", s.get("mean_one_way_turnover")],
        ["交易与组合", "累计交易成本率", s.get("cumulative_cost_rate")],
        ["交易与组合", "平均持仓品种数量", s.get("average_holding_count")],
        ["交易与组合", "调仓次数", s.get("rebalance_count")],
        ["交易与组合", "调仓成功率", success_rate],
        ["资金约束", "国债杠杆上限", s.get("bond_leverage_cap")],
        ["资金约束", "最低保证金率", s.get("bond_margin_floor")],
        ["资金约束", "平均收盘国债保证金率", s.get("mean_bond_margin_rate")],
        ["资金约束", "最低观察国债保证金率", s.get("minimum_bond_margin_rate")],
        ["资金约束", "最高观察国债杠杆", s.get("maximum_bond_leverage")],
        ["资金约束", "约束越界天数", s.get("margin_breach_days")],
        ["资金约束", "国债资金耗尽天数", s.get("bond_capital_depleted_days")],
        ["资金约束", "最大国债资金缺口", s.get("max_capital_shortfall")],
        ["资金约束", "资金约束状态", s.get("capital_constraint_status")],
        ["优化状态", "达到目标波动次数", s.get("target_achieved_count")],
        ["优化状态", "资金约束限制次数", s.get("cap_limited_count")],
    ]


def _style_performance(ws):
    percentages = {"累计收益率", "策略年化收益率", "基准年化收益率", "年化超额收益", "策略年化波动率",
                   "策略最大回撤", "年化跟踪误差", "超额最大回撤", "平均每日换仓比率（单边）", "累计交易成本率", "调仓成功率",
                   "基准年化波动率", "基准最大回撤", "最低保证金率", "平均收盘国债保证金率", "最低观察国债保证金率"}
    for row in ws.iter_rows(min_row=2):
        metric = row[1].value
        if metric in percentages:
            fmt = "0.00%"
        elif metric in {"年化换手率（单边）", "国债杠杆上限", "最高观察国债杠杆"}:
            fmt = '0.00"倍"'
        elif metric in {"交易日数量", "约束越界天数", "国债资金耗尽天数"}:
            fmt = '0"天"'
        elif metric in {"调仓次数", "达到目标波动次数", "资金约束限制次数"}:
            fmt = '0"次"'
        elif metric == "平均持仓品种数量":
            fmt = '0.00"种"'
        elif metric in {"初始净值", "期末净值", "Sharpe", "Sortino", "Calmar", "IR"}:
            fmt = "0.0000"
        else:
            fmt = "General"
        for cell in row[2:]:
            cell.number_format = fmt
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 34
    for index in range(3, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(index)].width = 28


def _parameter_rows(inputs, summary=None, mode="next_day_close"):
    s = summary or {}
    value = lambda key, default=None: inputs.get(key, s.get(key, default))
    rows = [
        ["基本信息", "执行价格", "下一交易日收盘价" if mode == "next_day_close" else mode],
        ["基本信息", "风险估计回看日数", value("lookback_days")],
        ["基本信息", "调仓频率", {"monthly": "月度", "quarterly": "季度", "semiannual": "半年度"}.get(value("rebalance_frequency"), value("rebalance_frequency"))],
        ["基本信息", "目标年化波动率", value("target_volatility", value("targets"))],
        ["基准设置", "基准名称", value("benchmark_name", "沪深300全收益指数")],
        ["基准设置", "基准代码", value("benchmark_code", "H00300.CSI")],
        ["基准设置", "基准相关指标", "共同基线和日期；不填充、不截短区间；超额净值=策略净值/基准净值"],
        ["基准设置", "超额指标口径", "主动日收益=策略日收益−基准日收益；年化超额收益=两者年化收益之差；超额回撤来自净值比值"],
        ["基准设置", "净值图起点", 1.0],
        ["基准设置", "净值标准化", "报告累计净值=组合净资产÷初始资金，从1起；账户净资产仍按初始资金记账；基线行不计入收益统计"],
        ["基准设置", "基线日期", s.get("baseline_date")],
        ["基准设置", "基线日期口径", "取回测首日前最近观察日，仅表示初始账户状态" if s.get("baseline_date_is_observed", True)
         else "无前序行情，使用前一自然日作为初始账户标记；该日不是行情观测"],
        ["组合约束", "目标仓位来源", "逐期优化文件的品种目标名义敞口；不重新优化、不归一化"],
        ["组合约束", "资产组", "股票、债券、黄金、其他商品"],
        ["组合约束", "组内分配", "继承当期优化名单和组内等名义敞口分配"],
        ["组合约束", "国债杠杆上限", value("bond_leverage_cap", value("leverage_caps"))],
        ["组合约束", "最低保证金率", value("bond_margin_floor")],
        ["组合约束", "调仓资金约束", "非债敞口+m_B×国债敞口=1；m_B不低于设定下限"],
        ["组合约束", "持有期间越界处理", "每日收盘记录交易前后资金；越界只标记，不触发额外减仓或强平"],
        ["账户结构", "账户模型", "国债资金=组合净资产−非债名义金额；保证金率=国债资金/国债名义金额"],
        ["账户结构", "资金耗尽显示", "国债资金不为正时杠杆为N/A；记录资金耗尽，结果不视为全程满足下限"],
        ["账户结构", "持仓单位", "加权价格序列的合成单位；不是实际期货合约手数"],
        ["交易设置", "初始资金", value("initial_nav", 1.0)],
        ["交易设置", "初始持仓", "回测开始时空仓；不承接开始日前的持仓"],
        ["交易设置", "成交时点", "信号日后下一观察日收盘调仓；旧仓取得执行日收益，新仓从后一观察日起计收益"],
        ["交易设置", "持仓期间", "固定合成单位数，名义敞口自然漂移，不每日再平衡"],
        ["交易设置", "样本末处理", "按最后收盘价估值，不强制平仓；无执行日期的信号不下单"],
        ["交易成本", "单边手续费（基点）", value("fee_bps", 0.0)],
        ["交易成本", "单边滑点（基点）", value("slippage_bps", 0.0)],
        ["交易成本", "费用计算", "买卖名义金额绝对值×费率；目标名义敞口以扣费后的净资产为分母"],
        ["交易成本", "单边换手率", "买卖名义金额绝对值合计÷调仓前扣费前净资产÷2"],
        ["交易成本", "年度及平均换手", "排除首次建仓日；平均值以其余观察日为分母，年度值按交易日数年化"],
        ["交易成本", "累计成本率", "1−逐日连乘(1−当日交易成本÷当日扣费前净资产)"],
        ["收益参数", "年化交易日", value("annualization", 252)],
        ["收益参数", "年化无风险利率", value("risk_free_rate_for_sharpe", 0.0)],
        ["收益参数", "无风险利率用途", "仅用于 Sharpe 和 Sortino，不计入账户收益"],
        ["收益参数", "滚动波动窗口", value("rolling_window", 63)],
        ["收益口径", "每日收益", "前日净资产加原持仓盈亏，减手续费及滑点；首次建仓成本计入当日收益"],
        ["收益口径", "年化收益", "(期末净资产÷初始资金)^(年化交易日÷观察日数)−1"],
        ["收益口径", "年化波动", "简单日收益样本标准差(ddof=1)×年化交易日平方根"],
        ["收益口径", "累计归因", "各组或品种累计盈亏减费用后除以初始资金；不单独复利每日贡献"],
        ["容量设置", "策略容量", "未估计；现有输入无成交额、实际合约及保证金数据，容量列留空"],
        ["数据口径", "行情", "同花顺 iFinD 加权期货日收盘价；用于价格序列敞口回测"],
        ["数据口径", "未模拟项目", "真实合约换月、合约乘数、整数手数、逐合约交易所保证金及强制平仓"],
        ["数据口径", "缺失价格", "持有或交易涉及的价格缺失时报错；不填充、不自动缩短区间"],
        ["输出口径", "未定义指标", "N/A；未估计容量以及基线收益为真正空白"],
    ]
    for key, label in (("price_source", "价格文件"), ("schedule_source", "调仓安排"), ("optimization_root", "优化结果目录"), ("benchmark_source", "基准文件")):
        if key in inputs:
            rows.append(["数据来源", label, inputs[key]])
    # 品种名称是持仓表的映射，不把整个字典挤进参数单元格。
    return rows


def _write_parameters(book, inputs, summary=None, mode="next_day_close", sheet="parameters"):
    ws = _table(book, sheet, ["类别", "参数 / 约束", "本项目设置"], _parameter_rows(inputs, summary, mode), max_width=50)
    for row in ws.iter_rows(min_row=2):
        if row[1].value in {"目标年化波动率", "年化无风险利率", "最低保证金率"}:
            row[2].number_format = "0.00%"
        if isinstance(row[2].value, str):
            row[2].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            width = sum(2 if ord(c) > 127 else 1 for c in row[2].value)
            ws.row_dimensions[row[0].row].height = max(18, math.ceil(width / 48) * 15)
    for column, width in (("A", 18), ("B", 40), ("C", 52)):
        ws.column_dimensions[column].width = width


def _annual_book(result):
    rows = []
    for r in _frame(result, "annual_returns").to_dict("records"):
        start, end = _date(r["first_date"], "annual first date"), _date(r["last_date"], "annual last date")
        if start.month > 1:
            label = f"{start.year}年{start:%m-%d}起"
        elif end.month < 12:
            label = f"{start.year}年截至{end:%m-%d}"
        else:
            label = f"{start.year}年"
        rows.append([label, _value(r, "return", "total_return", "net_return"), r.get("benchmark_return"), r.get("excess_return"),
                     r.get("annualized_volatility"), r.get("tracking_error"), r.get("sharpe_ratio"), r.get("sortino_ratio"), r.get("information_ratio"),
                     r.get("max_drawdown"), r.get("benchmark_max_drawdown"), r.get("excess_max_drawdown"), r.get("calmar_ratio"), r.get("annualized_one_way_turnover")])
    book = _new_book()
    ws = _table(book, "annual_metrics", ANNUAL_HEADERS, rows,
                percent={"策略收益", "基准收益", "超额收益", "年化波动", "跟踪误差", "策略最大回撤", "基准最大回撤", "超额最大回撤"},
                decimal={"Sharpe", "Sortino", "信息比率", "Calmar"}, max_width=24)
    for row in ws.iter_rows(min_row=2, min_col=14, max_col=14):
        row[0].number_format = '0.00"倍"'
    return book


def _performance_book(result, inputs, mode):
    book = _new_book()
    ws = _table(book, "performance", ["类别", "指标", "全天候策略"], _performance_rows(result), max_width=30)
    _style_performance(ws)
    _write_parameters(book, inputs, result["summary"], mode)
    return book


def _instrument_name(code, inputs):
    names = inputs.get("instrument_names", {})
    if not isinstance(names, dict):
        raise ValueError("instrument_names must be a dictionary")
    return names.get(code, code)


def _holdings_book(result, inputs):
    records = []
    frame = _frame(result, "instrument_daily", ("date", "instrument_code", "asset_class", "units_end", "exposure_end"))
    for r in frame.to_dict("records"):
        if r["units_end"] != 0:
            records.append([int(_date(r["date"], "holding date").strftime("%Y%m%d")), r["instrument_code"],
                            _instrument_name(r["instrument_code"], inputs), r["exposure_end"], GROUP_NAMES[r["asset_class"]]])
    book = _new_book()
    ws = _table(book, "holdings", ["日期", "品种代码", "品种名称", "名义敞口", "资产组"], records, decimal={"名义敞口"})
    for column, width in (("A", 14), ("B", 18), ("C", 38), ("D", 16), ("E", 18)):
        ws.column_dimensions[column].width = width
    return book


def _daily_records(result):
    daily = _frame(result, "daily_nav", ("date", "normalized_nav", "drawdown"))
    if daily.empty:
        raise ValueError("Daily NAV is empty")
    dates = pd.DatetimeIndex([_date(v, "daily date") for v in daily["date"]])
    if dates.has_duplicates or not dates.is_monotonic_increasing:
        raise ValueError("Daily dates must be unique and ascending")
    baseline = _date(result["summary"]["baseline_date"], "baseline_date")
    if baseline >= dates[0]:
        raise ValueError("Baseline date must precede the first return date")
    return daily.to_dict("records"), baseline


def _time_series_book(result):
    daily, baseline = _daily_records(result)
    group_frame = _frame(result, "group_daily", ("date", "asset_class", "exposure_end"))
    if group_frame.duplicated(["date", "asset_class"]).any():
        raise ValueError("Duplicate daily group exposures")
    group_map = {(pd.Timestamp(r["date"]), r["asset_class"]): r["exposure_end"] for r in group_frame.to_dict("records")}
    if any((pd.Timestamp(r["date"]), g) not in group_map for r in daily for g in GROUP_ORDER):
        raise ValueError("Missing daily group exposures")
    headers = ["日期", "策略日收益", "策略累计净值", "基准日收益", "基准累计净值", "主动日收益", "超额净值",
               "每日换仓比率（单边）", "策略可容纳规模（亿元）", "策略回撤", "毛敞口", "非债敞口",
               "股票敞口", "债券敞口", "黄金敞口", "其他商品敞口", "当日手续费", "当日滑点", "当日交易成本",
               "当日交易成本率", "累计交易成本率", "滚动年化波动率", "目标年化波动率", "调仓时预测年化波动率"]
    headers += ["基准回撤", "超额回撤", *FUNDING_COLUMNS.values()]
    first = [_BLANK] * len(headers)
    first[0], first[2] = baseline, 1.0
    first[4], first[6] = 1., 1.
    rows = [first]
    for r in daily:
        rows.append([r["date"], r["net_return"], r["normalized_nav"], r["benchmark_return"], r["benchmark_nav"], r["active_return"], r["relative_nav"],
                     r.get("one_way_turnover", r.get("turnover", 0.) / 2), _BLANK, r["drawdown"], r["gross_exposure"], r["non_bond_exposure"],
                     *[group_map[(pd.Timestamp(r["date"]), g)] for g in GROUP_ORDER],
                     r["fee"], r["slippage"], r["total_cost"], r.get("cost_rate"), r.get("cumulative_cost_rate"),
                     r.get("realized_volatility_rolling"), r.get("target_volatility"), r.get("forecast_volatility_at_rebalance"),
                     r["benchmark_drawdown"], r["relative_drawdown"], *_funding_values(r)])
    book = _new_book()
    percentages = set(headers) - {"日期", "策略累计净值", "基准累计净值", "超额净值", "策略可容纳规模（亿元）", "当日手续费", "当日滑点", "当日交易成本"}
    percentages = (percentages - set(FUNDING_COLUMNS.values())) | FUNDING_PERCENT
    ws = _table(book, "time_series", headers, rows, percent=percentages,
                decimal={"策略累计净值", "基准累计净值", "超额净值", "策略可容纳规模（亿元）", "当日手续费", "当日滑点", "当日交易成本"} | FUNDING_DECIMAL, max_width=34)
    for column, width in {"A": 15, "B": 20, "C": 20, "D": 28, "E": 20, "F": 20, "G": 20, "H": 28, "I": 30}.items():
        ws.column_dimensions[column].width = width
    for row in ws.iter_rows(min_row=2, min_col=9, max_col=9):
        row[0].number_format = "0.0000"
    return book


def _audit_table(book, name, frame, labels, percent=(), decimal=(), integer=()):
    if not isinstance(frame, pd.DataFrame):
        raise TypeError(f"{name} must be a DataFrame")
    keys = list(labels)
    rows = [[r.get(k) for k in keys] for r in frame.to_dict("records")]
    return _table(book, name, list(labels.values()), rows, percent=percent, decimal=decimal, integer=integer, max_width=36)


def _details_book(result, inputs, sources):
    book = _new_book()
    daily = _frame(result, "daily_nav").to_dict("records")
    by_date = {pd.Timestamp(r["date"]): r for r in daily}
    daily_rows = [[r["date"], r["nav"], r["normalized_nav"], r["gross_exposure"] * r["nav"], r["gross_exposure"], r["non_bond_exposure"],
                   r["trading_pnl"], r["fee"], r["slippage"], r["total_cost"],
                   "否" if _missing(r.get("rebalance_no")) else "是", r.get("held_signal_date"), r.get("rebalance_no"), *_funding_values(r)]
                  for r in daily]
    _table(book, "账户每日状态", ["日期", "组合净资产", "组合累计净值", "名义敞口金额", "毛敞口", "非债敞口",
           "当日交易盈亏", "手续费", "滑点", "交易成本", "当天是否调仓", "当前信号日", "调仓序号", *FUNDING_COLUMNS.values()], daily_rows,
           percent={"毛敞口", "非债敞口"} | FUNDING_PERCENT, decimal={"组合净资产", "组合累计净值", "名义敞口金额", "当日交易盈亏", "手续费", "滑点", "交易成本"} | FUNDING_DECIMAL, integer={"调仓序号"}, max_width=24)
    holding_rows, trade_rows = [], []
    for r in _frame(result, "instrument_daily").to_dict("records"):
        code, day = r["instrument_code"], pd.Timestamp(r["date"])
        name, group = _instrument_name(code, inputs), GROUP_NAMES[r["asset_class"]]
        signal = by_date[day].get("held_signal_date")
        holding_rows.append([r["date"], signal, code, name, group, r["units_end"], r["price"],
                             r["units_end"] * r["price"], r["exposure_end"], r["pnl"], r["net_return_contribution"],
                             "期末持有" if r["units_end"] != 0 else "当日已平仓"])
        if r["traded_notional"] != 0:
            trade_rows.append([r["date"], signal, code, name, group, "买入" if r["traded_notional"] > 0 else "卖出",
                               r["units_start"] * r["price"], r["units_end"] * r["price"], abs(r["traded_notional"]),
                               r["traded_notional"], r["fee"], r["slippage"], r["total_cost"]])
    _table(book, "账户持仓", ["日期", "信号日", "品种代码", "品种名称", "资产组", "序列单位数", "参考价格", "名义敞口金额",
           "对组合名义敞口", "当日盈亏", "净收益贡献", "持仓状态"], holding_rows, percent={"对组合名义敞口", "净收益贡献"},
           decimal={"序列单位数", "参考价格", "名义敞口金额", "当日盈亏"}, max_width=32)
    _table(book, "账户交易", ["成交日期", "信号日", "品种代码", "品种名称", "资产组", "交易方向", "交易前名义金额", "最新目标名义金额",
           "实际成交金额", "有符号成交金额", "手续费", "滑点", "交易成本"], trade_rows,
           decimal={"交易前名义金额", "最新目标名义金额", "实际成交金额", "有符号成交金额", "手续费", "滑点", "交易成本"}, max_width=32)
    allocation_rows = []
    for r in _frame(result, "rebalance_allocations").to_dict("records"):
        allocation_rows.append([r["date"], r["rebalance_no"], r["instrument_code"], GROUP_NAMES[r["asset_class"]],
                                r["target_exposure"], r["capital_weight"], r["target_notional"], r["capital_amount"], r["bond_margin_rate"]])
    _table(book, "调仓资金分配", ["日期", "调仓序号", "品种代码", "资产组", "目标名义敞口", "目标资金占比", "目标名义金额", "分配资金", "调仓目标国债保证金率"],
           allocation_rows, percent={"目标名义敞口", "目标资金占比", "调仓目标国债保证金率"}, decimal={"目标名义金额", "分配资金"}, integer={"调仓序号"})
    status_labels = {"target_achieved": "达到目标波动", "cap_limited": "资金约束限制", "not_loaded_not_required": "未执行，无需读取优化结果"}
    action_labels = {"executed": "已执行", "no_execution_date": "无后续执行日", "after_end_not_executed": "执行日在回测结束之后",
                     "before_start_not_inherited": "执行日在回测开始之前，未承接持仓"}
    signals = _frame(result, "signals").copy()
    if "optimization_status" in signals:
        signals["optimization_status"] = signals["optimization_status"].map(lambda v: status_labels.get(v, v))
    if "backtest_action" in signals:
        signals["backtest_action"] = signals["backtest_action"].map(lambda v: action_labels.get(v, v))
    _audit_table(book, "调仓信号", signals, {"rebalance_no": "调仓序号", "signal_date": "信号日", "execution_date": "执行日",
                 "optimization_status": "优化状态", "backtest_action": "回测处理", "target_volatility": "目标年化波动率",
                 "forecast_volatility": "预测年化波动率", "bond_leverage_cap": "国债杠杆上限", "bond_margin_floor": "最低保证金率",
                 "target_bond_margin_rate": "调仓目标国债保证金率", "source_path": "来源路径", "source_sha256": "来源SHA256"},
                 percent={"目标年化波动率", "预测年化波动率", "最低保证金率", "调仓目标国债保证金率"}, integer={"调仓序号"})
    attribution = _frame(result, "attribution").copy()
    if len(attribution):
        attribution["name"] = [GROUP_NAMES.get(r["name"], r["name"]) if r["level"] == "group" else
                               f"{r['name']} {_instrument_name(r['name'], inputs)}" for r in attribution.to_dict("records")]
        attribution["level"] = attribution["level"].map({"group": "资产组", "instrument": "品种"})
    _audit_table(book, "收益归因", attribution, {"level": "层级", "name": "资产", "trading_pnl": "累计交易盈亏", "fee": "累计手续费",
                 "slippage": "累计滑点", "total_cost": "累计交易成本", "net_pnl": "累计净盈亏", "contribution_to_total_return": "对累计收益的贡献"},
                 percent={"对累计收益的贡献"}, decimal={"累计交易盈亏", "累计手续费", "累计滑点", "累计交易成本", "累计净盈亏"})
    checks = _frame(result, "checks").copy()
    if "passed" in checks:
        checks["passed"] = checks["passed"].map(lambda v: "通过" if isinstance(v, (bool, np.bool_)) and v else "未通过")
    _audit_table(book, "校验", checks, {"check": "校验项目", "error": "误差", "tolerance": "容差", "passed": "是否通过", "detail": "说明"}, decimal={"误差", "容差"})
    source_rows = []
    for r in sources:
        source_rows.append([r.get("kind", r.get("source_type")), r.get("path", r.get("source_path")), r.get("sha256", r.get("source_sha256")),
                            r.get("source_price_sha256", _BLANK)])
    source_ws = _table(book, "来源", ["来源类型", "来源路径", "SHA256", "优化时价格SHA256"], source_rows, max_width=50)
    source_ws.column_dimensions["B"].width = 75
    source_ws.column_dimensions["C"].width = 67
    source_ws.column_dimensions["D"].width = 67
    for row in source_ws.iter_rows(min_row=2):
        row[1].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        source_ws.row_dimensions[row[0].row].height = max(18, math.ceil(len(str(row[1].value)) / 70) * 15)
    return book


def save_backtest_reports(output_dir, result, inputs, sources, *, mode="next_day_close", overwrite=False) -> list[Path]:
    """按 ETF 布局输出五个独立 Excel，含基准和整体国债资金；容量留空。"""
    paths = report_paths(output_dir, mode)
    for path in paths:
        _validate_destination(path, overwrite)
    books = []
    try:
        books.append(_annual_book(result))
        books.append(_performance_book(result, inputs, mode))
        books.append(_holdings_book(result, inputs))
        books.append(_time_series_book(result))
        books.append(_details_book(result, inputs, sources))
        for path, book in zip(paths, books):
            _save(book, path, overwrite)
    finally:
        for book in books:
            book.close()
    return paths


def save_comparison(path, results_by_scenario, inputs, overwrite=False):
    """每行一个目标波动率/杠杆上限组合，另列共同基准、净值和回撤。"""
    _validate_destination(Path(path), overwrite)
    if not results_by_scenario:
        raise ValueError("No backtest results to compare")
    columns = {"target_volatility": "目标年化波动率", "bond_leverage_cap": "国债杠杆上限", "bond_margin_floor": "最低保证金率",
               "annualized_return": "策略年化收益", "benchmark_annualized_return": "基准年化收益", "annualized_excess_return": "年化超额收益",
               "annualized_volatility": "策略年化波动率", "max_drawdown": "最大回撤", "sharpe_ratio": "Sharpe",
               "tracking_error": "跟踪误差", "information_ratio": "IR", "annualized_one_way_turnover": "年化单边换手率",
               "cumulative_cost_rate": "累计交易成本率", "minimum_bond_margin_rate": "最低观察保证金率",
               "maximum_bond_leverage": "最高观察杠杆", "margin_breach_days": "约束越界天数",
               "bond_capital_depleted_days": "国债资金耗尽天数", "target_achieved_count": "达到目标波动次数",
               "cap_limited_count": "资金约束限制次数", "capital_constraint_status": "资金约束状态"}
    records, nav_series, drawdown_series = [], [], []
    labels = []
    common_dates, common_benchmark = None, None
    for (target, limit), result in sorted(results_by_scenario.items()):
        target = _number(target, "target_volatility", True)
        limit = _number(limit, "bond_leverage_cap", True)
        _same_number(result["summary"]["target_volatility"], target, "scenario target")
        _same_number(result["summary"]["bond_leverage_cap"], limit, "scenario leverage cap")
        label = f"目标{target * 100:g}%_国债上限{limit:g}倍"
        if label in labels:
            raise ValueError("Duplicate target labels")
        labels.append(label)
        records.append([result["summary"].get(key) for key in columns])
        daily, baseline = _daily_records(result)
        dates = pd.DatetimeIndex([baseline, *[r["date"] for r in daily]])
        benchmark = np.array([1., *[r["benchmark_nav"] for r in daily]])
        if common_dates is None:
            common_dates, common_benchmark = dates, benchmark
        elif not common_dates.equals(dates) or not np.array_equal(common_benchmark, benchmark):
            raise ValueError("Scenario comparisons require identical dates and benchmark paths")
        nav_series.append(pd.Series([1., *[r["normalized_nav"] for r in daily]], index=dates, name=label))
        drawdown_series.append(pd.Series([0., *[r["drawdown"] for r in daily]], index=dates, name=label))
    nav_series.append(pd.Series(common_benchmark, index=common_dates, name="沪深300全收益"))
    drawdown_series.append(pd.Series(common_benchmark / np.maximum.accumulate(common_benchmark) - 1, index=common_dates, name="沪深300全收益"))
    book = _new_book()
    try:
        _table(book, "指标对比", list(columns.values()), records,
               percent={"目标年化波动率", "最低保证金率", "策略年化收益", "基准年化收益", "年化超额收益", "策略年化波动率",
                        "最大回撤", "跟踪误差", "累计交易成本率", "最低观察保证金率"},
               decimal={"国债杠杆上限", "Sharpe", "IR", "年化单边换手率", "最高观察杠杆"},
               integer={"约束越界天数", "国债资金耗尽天数", "达到目标波动次数", "资金约束限制次数"}, max_width=30)
        for name, series, percentages in (("净值对比", nav_series, False), ("回撤对比", drawdown_series, True)):
            frame = pd.concat(series, axis=1).sort_index().rename_axis("日期").reset_index()
            records = [[_BLANK if _missing(v) else v for v in row] for row in frame.itertuples(index=False, name=None)]
            value_columns = list(frame.columns[1:])
            _table(book, name, frame.columns, records, percent=value_columns if percentages else (), decimal=() if percentages else value_columns, max_width=24)
        _write_parameters(book, {**inputs, "targets": tuple(sorted({key[0] for key in results_by_scenario})),
                                 "leverage_caps": tuple(sorted({key[1] for key in results_by_scenario}, reverse=True))}, sheet="参数")
        _save(book, Path(path), overwrite)
    except BaseException:
        book.close()
        raise
