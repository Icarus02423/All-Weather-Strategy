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


ANNUAL_HEADERS = ["年份", "策略收益", "基准收益", "超额收益", "年化波动", "跟踪误差", "Sharpe", "Sortino",
                  "信息比率", "策略最大回撤", "基准最大回撤", "超额最大回撤", "Calmar", "年化换手率（单边）"]






















































