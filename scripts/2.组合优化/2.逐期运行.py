"""读取已有逐期风险 Excel，调用同目录的“3.优化求解.py”，输出逐期优化 Excel。

Python 3.9+；使用项目已有 numpy、pandas、openpyxl，不生成 JSON 文件。
直接运行处理顶部所选方案；导入本文件不会读取、计算或保存数据。

已有风险工作簿是四组 Q 和当期名单的来源。旧工作簿未保存品种 Sigma，
因此只读原价格，严格按该工作簿的窗口和名单补算 A、Sigma，并与原来的
ClassReturns、Q 交叉核对。优化始终使用 Excel 中的 Q，不替换、不修补。
准备表只用于核对原安排，不重新制定新品种纳入规则。输出仍为 Excel。

目录示例：outputs/optimization/Lookback_Quarterly_Rebalance_Monthly/
TargetVol_05pct/001.xlsx。001 是原调仓序号，不是运行次数。
默认不覆盖；失败先保存诊断工作簿再停止。测试请显式使用独立输出目录。
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json  # 仅读取已有风险 Excel 属性中的元数据；不创建或读取 .json 文件。
import math
import os
import re
import sys
import tempfile
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Optional

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ---------- 日常运行设置：在这里选择要读的风险方案及日期 ----------
PROJECT_ROOT = Path(__file__).resolve().parents[2]
LOOKBACK_PERIOD = "quarterly"        # quarterly / semiannual / annual
REBALANCE_FREQUENCY = "monthly"      # monthly / quarterly / semiannual
START_SIGNAL_DATE = "2014-12-31"
END_SIGNAL_DATE = "2026-09-30"
TARGET_VOLATILITIES = (0.05, 0.10, 0.15)
RISK_BUDGETS = {"equity": 1 / 3, "bonds": 1 / 3,
                "gold": 1 / 6, "other_commodities": 1 / 6}
NON_BOND_GROUPS = ("equity", "gold", "other_commodities")  # 合计上限固定为1
OVERWRITE_OUTPUTS = False
FAILURE_POLICY = "stop"              # stop / continue；失败从不变成有效仓位

# ---------- 数值求解设置：逐次传入求解器，不另存一份配置文件 ----------
INITIAL_POINTS = None               # None 使用求解器公开说明的三个正初值
MAX_ITERATIONS = 100
MAX_BACKTRACKS = 60
KKT_TOLERANCE = 1e-13
BUDGET_TOLERANCE = 1e-8
CAP_TOLERANCE = 1e-10
VOLATILITY_TOLERANCE = 1e-10
TRACE_LEVEL = "full"                # full 保存逐轮矩阵与试探步；standard 省矩阵
INPUT_ABSOLUTE_TOLERANCE = 1e-12
INPUT_RELATIVE_TOLERANCE = 1e-10

GROUP_ORDER = ("equity", "bonds", "gold", "other_commodities")  # S,B,G,C
GROUP_SYMBOLS = ("S", "B", "G", "C")
LOOKBACK_WINDOWS = {"quarterly": 63, "semiannual": 126, "annual": 252}
FREQUENCY_NAMES = {"monthly": "Monthly", "quarterly": "Quarterly", "semiannual": "Semiannual"}
PERIOD_NAMES = {"quarterly": "Quarterly", "semiannual": "Semiannual", "annual": "Annual"}
RUNNER_VERSION = "new_objective_excel_v1"
SUCCESS_STATUSES = {"target_achieved", "cap_limited"}


class BatchFailure(RuntimeError):
    """本次任务失败，诊断文件已保存；不是原优化问题不可行的结论。"""


def _digest(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _load_module(path: Path, name: str):
    """保留用户的数字中文文件名，按路径导入；不在项目写 __pycache__。"""
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"Required Python module is missing: {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    old_flag = sys.dont_write_bytecode
    previous = sys.modules.get(name)
    sys.dont_write_bytecode = True
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except BaseException:
        if previous is None:
            sys.modules.pop(name, None)
        else:
            sys.modules[name] = previous
        raise
    finally:
        sys.dont_write_bytecode = old_flag
    return module


def _date(value: Any, field: str, nullable: bool = False):
    if value is None or value == "":
        if nullable:
            return None
        raise ValueError(f"Missing {field}")
    if isinstance(value, (bool, int, float, np.number)):
        raise ValueError(f"{field} must be a date, not a numeric value")
    if isinstance(value, str) and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        raise ValueError(f"{field} must use YYYY-MM-DD")
    value = pd.Timestamp(value)
    if pd.isna(value) or value.tzinfo is not None or value != value.normalize():
        raise ValueError(f"Invalid {field}")
    return value


def _number(value: Any, field: str) -> float:
    if isinstance(value, (bool, str)) or not isinstance(value, (int, float, np.number)):
        raise ValueError(f"{field} must be numeric")
    value = float(value)
    if not math.isfinite(value):
        raise ValueError(f"{field} must be finite")
    return value


def _section(rows: list[tuple], title: str, headers: tuple[str, ...], *, exact_headers: bool = False) -> list[dict]:
    """按现有工作簿的区块标题及列名读取，不依赖区块所在的固定行号。"""
    found = [(i, j) for i, row in enumerate(rows) for j, value in enumerate(row) if value == title]
    if len(found) != 1:
        raise ValueError(f"Expected exactly one {title} section")
    title_row, start_column = found[0]
    i = title_row + 1
    # 同一行可能有并列区块；只解析从标题所在列开始的连续表头。
    header = []
    if i < len(rows):
        for value in rows[i][start_column:]:
            if value in (None, ""):
                break
            header.append(value)
    if any(header.count(h) != 1 for h in headers):
        raise ValueError(f"Invalid headers in {title}")
    if exact_headers and (len(header) != len(headers) or set(header) != set(headers)):
        raise ValueError(f"Unexpected columns in {title}")
    positions = {h: header.index(h) for h in headers}
    result = []
    for row in rows[i + 1:]:
        if len(row) <= start_column or row[start_column] in (None, ""):
            break
        result.append({h: row[start_column + j] if start_column + j < len(row) else None
                       for h, j in positions.items()})
    if not result:
        raise ValueError(f"Empty {title} section")
    return result


def read_risk_workbook(path: Path) -> dict:
    """读取现有单期风险工作簿，并统一成文稿的 S,B,G,C 顺序。"""
    path = Path(path)
    original_hash = _digest(path)
    # 现有风险输出是数值快照；拒绝公式，避免把不明来源的旧缓存当输入。
    formula_book = load_workbook(path, read_only=True, data_only=False)
    try:
        if any(cell.data_type == "f" for sheet in formula_book for row in sheet for cell in row):
            raise ValueError("Risk input must contain saved numeric observations, not formulas")
    finally:
        formula_book.close()
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        if len(book.worksheets) != 1:
            raise ValueError("Expected the existing one-rebalance-per-workbook layout")
        rows = list(book.active.iter_rows(values_only=True))
        markers = [i for i, row in enumerate(rows) if row and row[0] == "AnnualCovariance"]
        if len(markers) != 1:
            raise ValueError("Missing or duplicate AnnualCovariance section")
        fields = {}
        for row in rows[:markers[0]]:
            for j in (0, 3, 6):
                if len(row) > j + 1 and isinstance(row[j], str) and re.fullmatch(r"[a-z]+(?:_[a-z]+)+", row[j]):
                    if row[j] in fields:
                        raise ValueError(f"Duplicate period field: {row[j]}")
                    fields[row[j]] = row[j + 1]
        required = {"rebalance_no", "signal_date", "execution_date", "lookback_period", "lookback_days",
                    "rebalance_frequency", "annualization_days", "covariance_ddof", "return_method",
                    "price_start", "price_end", "return_start", "return_end", "n_prices", "n_returns",
                    "risk_status", "execution_status"}
        if required - fields.keys():
            raise ValueError(f"Missing period fields: {sorted(required - fields.keys())}")
        for key in ("rebalance_no", "lookback_days", "annualization_days", "covariance_ddof", "n_prices", "n_returns"):
            number = _number(fields[key], key)
            if not number.is_integer() or number < 1:
                raise ValueError(f"Invalid integer {key}")
            fields[key] = int(number)
        for key in ("signal_date", "price_start", "price_end", "return_start", "return_end"):
            fields[key] = _date(fields[key], key)
        fields["execution_date"] = _date(fields["execution_date"], "execution_date", True)
        if fields["return_method"] != "simple" or fields["covariance_ddof"] != 1:
            raise ValueError("Only the agreed simple returns and covariance ddof=1 are supported")
        if fields["risk_status"] != "ok":
            raise ValueError("Source workbook is not a successful risk estimate")
        if fields["lookback_period"] not in LOOKBACK_WINDOWS or fields["lookback_days"] != LOOKBACK_WINDOWS[fields["lookback_period"]]:
            raise ValueError("Lookback name and day count disagree")
        if fields["rebalance_frequency"] not in FREQUENCY_NAMES:
            raise ValueError("Unknown rebalance frequency")
        if not re.fullmatch(r"[0-9]+", path.stem) or int(path.stem) != fields["rebalance_no"]:
            raise ValueError("Filename and original rebalance number disagree")
        if book.active.title != f"Rebalance_{fields['rebalance_no']:03d}":
            raise ValueError("Worksheet name and original rebalance number disagree")
        covariance_rows = _section(rows, "AnnualCovariance", ("asset_class", *GROUP_ORDER))
        if len(covariance_rows) != 4 or Counter(r["asset_class"] for r in covariance_rows) != Counter(GROUP_ORDER):
            raise ValueError("Covariance must contain exactly the four labelled groups")
        by_group = {r["asset_class"]: r for r in covariance_rows}
        Q = np.array([[_number(by_group[g][other], "Q") for other in GROUP_ORDER] for g in GROUP_ORDER])
        if not np.array_equal(Q, Q.T):
            raise ValueError("Excel Q is not exactly symmetric; it will not be repaired")
        try:
            np.linalg.cholesky(Q)
        except np.linalg.LinAlgError as exc:
            raise ValueError("Excel Q is not positive definite, as required by the current method") from exc
        allocations = {g: {} for g in GROUP_ORDER}
        codes = set()
        for row in _section(rows, "Allocations", ("asset_class", "instrument_code", "within_class_weight")):
            group, code = row["asset_class"], row["instrument_code"]
            if group not in allocations or not isinstance(code, str) or not code or code != code.strip() or code in codes:
                raise ValueError("Invalid/duplicate group or instrument in Allocations")
            codes.add(code)
            allocations[group][code] = _number(row["within_class_weight"], "within_class_weight")
        for group, members in allocations.items():
            if not members or any(abs(w - 1 / len(members)) > 1e-12 for w in members.values()):
                raise ValueError(f"{group} must be nonempty and equally weighted")
        observations = _section(rows, "ClassReturns", ("return_date", *GROUP_ORDER))
        days = pd.DatetimeIndex([_date(r["return_date"], "return_date") for r in observations])
        if not days.is_unique or not days.is_monotonic_increasing:
            raise ValueError("Class return dates must be unique and ascending")
        group_returns = np.array([[_number(r[g], "class_return") for g in GROUP_ORDER] for r in observations])
        diagnostics = {r["field"]: r["value"] for r in _section(rows, "Diagnostics", ("field", "value"))}
        instrument_codes = [code for group in GROUP_ORDER for code in allocations[group]]
        instrument_rows = _section(rows, "InstrumentAnnualCovariance", ("instrument_code", *instrument_codes), exact_headers=True)
        if Counter(r["instrument_code"] for r in instrument_rows) != Counter(instrument_codes):
            raise ValueError("Instrument covariance rows must match Allocations exactly")
        by_code = {r["instrument_code"]: r for r in instrument_rows}
        Sigma = np.array([[_number(by_code[code][other], "Sigma") for other in instrument_codes]
                          for code in instrument_codes])
        if not np.array_equal(Sigma, Sigma.T) or np.any(np.diag(Sigma) < 0):
            raise ValueError("Instrument covariance must be symmetric with nonnegative variances")
        # 品种矩阵可以奇异；不增加正定要求，也不修补矩阵。
        instrument_observations = _section(rows, "InstrumentReturns", ("return_date", *instrument_codes), exact_headers=True)
        instrument_days = pd.DatetimeIndex([_date(r["return_date"], "instrument return date") for r in instrument_observations])
        if not instrument_days.equals(days):
            raise ValueError("InstrumentReturns dates must match ClassReturns exactly")
        instrument_returns = np.array([[_number(r[c], "instrument return") for c in instrument_codes]
                                       for r in instrument_observations])
        volatilities = {}
        for title, label, labels in (("AnnualVolatility", "asset_class", GROUP_ORDER),
                                     ("InstrumentAnnualVolatility", "instrument_code", instrument_codes)):
            entries = _section(rows, title, (label, "annual_volatility"), exact_headers=True)
            if Counter(r[label] for r in entries) != Counter(labels):
                raise ValueError(f"Unexpected or duplicate labels in {title}")
            by_label = {r[label]: _number(r["annual_volatility"], title) for r in entries}
            volatilities[title] = np.array([by_label[name] for name in labels])
        instrument_diagnostics = {r["field"]: r["value"] for r in
                                  _section(rows, "InstrumentDiagnostics", ("field", "value"), exact_headers=True)}
        description = book.properties.description
        try:
            metadata = json.loads(description) if description else {}
        except (ValueError, TypeError) as exc:
            raise ValueError("Invalid metadata embedded in the existing risk workbook") from exc
        if not isinstance(metadata, dict):
            raise ValueError("Invalid risk metadata")
        for key in ("rebalance_no", "lookback_period", "lookback_days", "rebalance_frequency", "annualization_days", "covariance_ddof", "return_method"):
            if key in metadata and metadata[key] != fields[key]:
                raise ValueError(f"Workbook metadata contradicts visible field {key}")
        if "signal_date" in metadata and _date(metadata["signal_date"], "metadata.signal_date") != fields["signal_date"]:
            raise ValueError("Metadata signal_date contradicts the visible date")
    finally:
        book.close()
    if _digest(path) != original_hash:
        raise ValueError("Risk workbook changed while being read")
    return {"fields": fields, "Q": Q, "allocations": allocations,
            "return_dates": days, "class_returns": group_returns,
            "Sigma": Sigma, "instrument_codes": instrument_codes, "instrument_returns": instrument_returns,
            "annual_volatility": volatilities["AnnualVolatility"],
            "instrument_annual_volatility": volatilities["InstrumentAnnualVolatility"],
            "source_instrument_diagnostics": instrument_diagnostics,
            "source_diagnostics": diagnostics, "metadata": metadata,
            "source_path": path.resolve(), "source_sha256": original_hash}


def prepare_period(risk: dict, schedule_row: dict, prices: pd.DataFrame,
                   price_hash: str, schedule_hash: str) -> dict:
    """只读补算旧 Excel 未存的 A、Sigma；核对通过后仍使用原 Excel Q。"""
    fields = risk["fields"]
    for key, other in (("rebalance_no", "rebalance_no"), ("signal_date", "signal_date"),
                       ("execution_date", "execution_date"), ("lookback_days", "lookback"),
                       ("rebalance_frequency", "frequency")):
        if fields[key] != schedule_row[other]:
            raise ValueError(f"Risk workbook and preparation schedule disagree on {key}")
    allocations = risk["allocations"]
    for group in GROUP_ORDER:
        planned = schedule_row["allocations"][group]
        if set(allocations[group]) != set(planned) or any(abs(allocations[group][c] - planned[c]) > 1e-12 for c in planned):
            raise ValueError(f"Risk workbook and schedule disagree on {group} membership/weights")
    signal, length = fields["signal_date"], fields["lookback_days"]
    if signal not in prices.index:
        raise ValueError("Signal date is absent from the observed price calendar")
    end = prices.index.get_loc(signal)
    if end < length:
        raise ValueError("Insufficient price history; the window will not be extended or filled")
    codes = risk["instrument_codes"]
    if any(c not in prices.columns for c in codes):
        raise ValueError("A selected instrument is absent from the price table")
    window = prices.loc[:, codes].iloc[end - length:end + 1]
    values = window.to_numpy(dtype=float)
    if not np.isfinite(values).all() or np.any(values <= 0):
        raise ValueError("Selected price window contains missing, nonfinite or nonpositive prices")
    with np.errstate(over="raise", divide="raise", invalid="raise"):
        price_returns = values[1:] / values[:-1] - 1
    expected_dates = window.index[1:]
    if not risk["return_dates"].equals(expected_dates):
        raise ValueError("ClassReturns dates do not match the exact L-return price window")
    expected_fields = {"price_start": window.index[0], "price_end": signal,
                       "return_start": expected_dates[0], "return_end": signal,
                       "n_prices": length + 1, "n_returns": length}
    if any(fields[k] != v for k, v in expected_fields.items()):
        raise ValueError("Recorded sample endpoints/counts disagree with the selected window")
    execution = fields["execution_date"]
    if execution is not None:
        if end + 1 >= len(prices) or prices.index[end + 1] != execution:
            raise ValueError("Execution date is not the next observed price date")
        if fields["execution_status"] != "scheduled":
            raise ValueError("Execution date and status disagree")
    elif fields["execution_status"] != "no_execution_date":
        raise ValueError("Missing execution date must be explicitly marked no_execution_date")
    A = np.zeros((len(codes), 4))
    for i, code in enumerate(codes):
        for j, group in enumerate(GROUP_ORDER):
            A[i, j] = allocations[group].get(code, 0)
    factor = fields["annualization_days"]
    returns, Sigma = risk["instrument_returns"], risk["Sigma"]
    group_returns = returns @ A
    direct_Q = np.cov(risk["class_returns"], rowvar=False, ddof=1) * factor
    mapped_Q = A.T @ Sigma @ A
    checks = []
    for name, actual, expected in (
            ("saved_instrument_returns_match_prices", returns, price_returns),
            ("saved_Sigma_matches_instrument_returns", Sigma, np.cov(returns, rowvar=False, ddof=1) * factor),
            ("saved_class_returns_match_prices", price_returns @ A, risk["class_returns"]),
            ("saved_class_returns_match_instruments", group_returns, risk["class_returns"]),
            ("saved_Q_matches_class_returns", direct_Q, risk["Q"]),
            ("saved_Q_matches_A_Sigma_A", mapped_Q, risk["Q"]),
            ("saved_class_volatility_matches_Q", risk["annual_volatility"], np.sqrt(np.diag(risk["Q"]))),
            ("saved_instrument_volatility_matches_Sigma", risk["instrument_annual_volatility"], np.sqrt(np.diag(Sigma)))):
        difference = float(np.max(np.abs(actual - expected)))
        passed = bool(np.allclose(actual, expected, rtol=INPUT_RELATIVE_TOLERANCE, atol=INPUT_ABSOLUTE_TOLERANCE))
        checks.append({"check": name, "error": difference, "tolerance": INPUT_ABSOLUTE_TOLERANCE,
                       "passed": passed, "detail": f"Also relative tolerance {INPUT_RELATIVE_TOLERANCE:g}"})
        if not passed:
            raise ValueError(f"{name} failed: maximum absolute error {difference:.17g}")
    metadata = risk["metadata"]
    return {**risk, "A": A, "Sigma": Sigma, "instrument_codes": codes,
            "instrument_returns": returns, "price_dates": window.index, "price_values": values,
            "reconstructed_class_returns": group_returns, "source_checks": checks,
            "schedule_source_row": schedule_row["source_row"],
            "price_sha256": price_hash, "schedule_sha256": schedule_hash,
            "source_price_hash_matches": metadata.get("price_sha256") == price_hash,
            "source_schedule_hash_matches": metadata.get("schedule_sha256") == schedule_hash,
            "window_sha256": hashlib.sha256(window.to_csv(float_format="%.17g").encode()).hexdigest()}


def complete_result(risk: dict, solution: dict, b: np.ndarray, h: np.ndarray,
                    target: float, solver_options: dict) -> dict:
    """独立从输入矩阵复算最终指标；禁止截断、归一化或修改求得的仓位。"""
    result = {"solution": solution, "status": solution["status"], "checks": list(risk["source_checks"]),
              "groups": [], "instruments": []}
    if solution["status"] not in SUCCESS_STATUSES:
        return result
    x = np.asarray(solution["x"], dtype=float)
    Q, A, Sigma = risk["Q"], risk["A"], risk["Sigma"]
    qx = Q @ x
    variance = float(x @ qx)
    volatility = float(np.sqrt(variance))
    group_mrc = qx / volatility
    group_rc = x * group_mrc
    group_shares = group_rc / volatility
    w = A @ x
    instrument_mrc = Sigma @ w / volatility
    instrument_rc = w * instrument_mrc
    cap = float(h @ x)
    cap_multiplier = float(solution["cap_multiplier"])
    budget_jacobian = np.array([np.eye(4)[g] * qx[g] + x[g] * Q[g] - 2 * b[g] * qx for g in range(3)])
    gradient = 2 * (1 - target / volatility) * qx
    stationarity = gradient + budget_jacobian.T @ solution["budget_multipliers"] + cap_multiplier * h

    def check(name, error, tolerance, passed=None, detail=""):
        finite = math.isfinite(float(error)) and math.isfinite(float(tolerance))
        result["checks"].append({"check": name, "error": float(error), "tolerance": float(tolerance),
                                 "passed": bool(finite and (error <= tolerance if passed is None else passed)), "detail": detail})
    check("positive_group_exposures", float(np.min(x)), 0, bool(np.all(x > 0)), "Strictly greater than zero")
    check("all_four_risk_shares", np.max(np.abs(group_shares - b)), solver_options["budget_tol"])
    check("non_bond_cap", max(0, cap - 1), solver_options["cap_tol"])
    check("stationarity", np.max(np.abs(stationarity)), solver_options["residual_tol"])
    check("nonnegative_cap_multiplier", max(0, -cap_multiplier), solver_options["residual_tol"])
    check("complementarity", abs(cap_multiplier * (cap - 1)), solver_options["cap_tol"])
    consistency_tol = INPUT_ABSOLUTE_TOLERANCE + INPUT_RELATIVE_TOLERANCE * max(volatility, variance)
    check("variance_groups_equals_instruments", abs(variance - w @ Sigma @ w), consistency_tol)
    check("group_risk_contributions_sum", abs(group_rc.sum() - volatility), consistency_tol)
    check("instrument_risk_contributions_sum", abs(instrument_rc.sum() - volatility), consistency_tol)
    for j, group in enumerate(GROUP_ORDER):
        mask = A[:, j] > 0
        check(f"{group}_instrument_contributions", abs(instrument_rc[mask].sum() - group_rc[j]), consistency_tol)
        result["groups"].append((group, x[j], b[j], group_shares[j], group_shares[j] - b[j],
                                 group_mrc[j], group_rc[j], x[j] * qx[j], float(solution["budget_multipliers"][j]) if j < 3 else None))
    if solution["status"] == "target_achieved":
        check("target_achieved", abs(volatility - target), solver_options["vol_tol"])
    else:
        check("active_cap_equality", abs(cap - 1), solver_options["cap_tol"])
        check("cap_limited_below_target", volatility - target, 0, volatility < target)
        check("cap_limited_positive_multiplier", cap_multiplier, 0, cap_multiplier > 0, "Strictly positive")
    for i, code in enumerate(risk["instrument_codes"]):
        j = int(np.flatnonzero(A[i])[0])
        result["instruments"].append((code, GROUP_ORDER[j], True, A[i, j], w[i], instrument_mrc[i],
                                      instrument_rc[i], instrument_rc[i] / volatility))
    if any(not r["passed"] for r in result["checks"]):
        result["status"] = "postcheck_failed"
        result["groups"], result["instruments"] = [], []
        result["error"] = "Independent validation failed; candidate positions are not valid target positions"
    result.update({"volatility": volatility, "objective": (volatility - target) ** 2,
                   "cap_exposure": cap, "gross_exposure": float(x.sum()), "cap_multiplier": cap_multiplier})
    return result


def _cell(value: Any):
    if value is None:
        return None
    if isinstance(value, np.generic):
        value = value.item()
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, (str, bool, int, datetime, date)):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else str(value)
    return str(value)


def _table(book: Workbook, title: str, headers: tuple, records: list[tuple]):
    ws = book.create_sheet(title)
    ws.append(headers)
    for row in records:
        ws.append([_cell(v) for v in row])
    ws.freeze_panes = "A2"
    if records:
        ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="334B5F")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Calibri", size=11)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if isinstance(cell.value, str):
                cell.data_type = "s"  # 路径/错误信息是文字，不能作为 Excel 公式执行。
            elif isinstance(cell.value, (datetime, date)):
                cell.number_format = "yyyy-mm-dd"
            elif isinstance(cell.value, float):
                cell.number_format = "0.000000000000E+00" if (0 < abs(cell.value) < 1e-5) else "0.000000000000"
    for i, header in enumerate(headers, 1):
        width = min(42, max(17, len(str(header)) + 2))
        if header in {"detail", "value", "reason", "source_path"}:
            width = 65
        ws.column_dimensions[get_column_letter(i)].width = width
    for row in ws.iter_rows(min_row=2):
        lines = max([1] + [sum(max(1, math.ceil(len(part) / max(10, ws.column_dimensions[cell.column_letter].width - 2)))
                              for part in cell.value.split("\n"))
                             for cell in row if isinstance(cell.value, str)])
        ws.row_dimensions[row[0].row].height = min(180, 15 * lines)
    return ws


def _matrix_rows(label: str, matrix: Any, row_labels: list, col_labels: list,
                 prefix: tuple = ()) -> list[tuple]:
    matrix = np.asarray(matrix)
    if matrix.shape != (len(row_labels), len(col_labels)):
        raise ValueError(f"Unexpected matrix shape while exporting {label}")
    return [(*prefix, label, r, c, float(matrix[i, j]))
            for i, r in enumerate(row_labels) for j, c in enumerate(col_labels)]


def save_workbook(path: Path, record: dict, risk: Optional[dict], result: dict,
                  b: np.ndarray, h: np.ndarray, target: float, solver_options: dict,
                  sources: dict, overwrite: bool = False) -> None:
    """保存一个时期、一个目标的结果；失败仍保留输入及已有求解轨迹。"""
    path = Path(path)
    if path.is_symlink() or (path.exists() and not overwrite):
        raise FileExistsError(f"Output exists or is a symbolic link: {path}")
    book = Workbook()
    book.remove(book.active)
    solution = result.get("solution", {})
    overview = [("status", result["status"]), ("rebalance_no", record["rebalance_no"]),
                ("signal_date", record["signal_date"]), ("execution_date", record["execution_date"]),
                ("execution_status", "scheduled" if record["execution_date"] is not None else "no_execution_date"),
                ("target_volatility", target), ("actual_volatility", result.get("volatility")),
                ("volatility_gap", result.get("volatility", target) - target if "volatility" in result else None),
                ("objective", result.get("objective")), ("non_bond_exposure", result.get("cap_exposure")),
                ("gross_exposure", result.get("gross_exposure")), ("cap_multiplier", result.get("cap_multiplier")),
                ("error", result.get("error", solution.get("reason", ""))),
                ("attempt_count", len(solution.get("attempts", []))),
                ("global_optimality_basis", solution.get("global_optimality_basis", ""))]
    _table(book, "Overview", ("field", "value"), overview)
    _table(book, "GroupResults", ("asset_class", "target_exposure", "budget", "risk_share", "budget_error",
                                   "marginal_risk", "risk_contribution", "variance_contribution", "budget_multiplier"), result.get("groups", []))
    _table(book, "InstrumentResults", ("instrument_code", "asset_class", "included", "within_class_weight",
                                        "target_exposure", "marginal_risk", "risk_contribution", "risk_share"), result.get("instruments", []))
    attempt_rows, iteration_rows, trial_rows, matrix_rows = [], [], [], []
    for index, attempt in enumerate(solution.get("attempts", []), 1):
        branch = "active_cap" if attempt["active_cap"] else "inactive_cap"
        attempt_rows.append((index, branch, attempt["status"], attempt.get("accepted", False),
                             attempt.get("initial_point_source", ""), *attempt["initial_x"],
                             attempt.get("acceptance_reason") or attempt.get("rejection_reason") or attempt.get("reason", ""),
                             attempt.get("detail", "")))
        for row in attempt.get("history", []):
            multipliers = list(row["multipliers"])
            if len(multipliers) == 3:
                multipliers.append(0.0)
            x = np.asarray(row["x"])
            Q = risk["Q"] if risk else None
            sigma = math.sqrt(float(x @ Q @ x)) if Q is not None else None
            shares = x * (Q @ x) / (sigma * sigma) if sigma is not None else [None] * 4
            risk_eq = x * (Q @ x) - b * sigma**2 if sigma is not None else [None] * 4
            iteration_rows.append((index, branch, row["iteration"], *x, *multipliers,
                                   row.get("volatility", sigma), row.get("objective", (sigma - target)**2 if sigma else None),
                                   row["cap_exposure"], *shares, *risk_eq,
                                   *row.get("stationarity", [None] * 4), row["residual_max"],
                                   row.get("step_size"), row.get("backtracks"), row.get("linear_solve_residual_max"),
                                   row.get("line_search_evaluations"), row.get("update_accepted"), row.get("termination_reason")))
            for trial_no, trial in enumerate(row.get("line_search_trials", []), 1):
                trial_rows.append((index, branch, row["iteration"], trial_no,
                                   trial.get("step_size"), trial.get("accepted"), trial.get("reason"),
                                   trial.get("residual_norm"), trial.get("required_residual_norm_max"),
                                   *trial.get("candidate_x", [None] * 4),
                                   *(list(trial.get("candidate_multipliers", [])) + [None] * 4)[:4], trial.get("detail", "")))
            for key, value in row.items():
                if key not in {"objective_gradient", "objective_hessian", "constraint_jacobian", "lagrangian_hessian",
                               "kkt_matrix", "residual", "rhs", "direction", "qx"} or value is None:
                    continue
                array = np.asarray(value)
                if array.ndim == 1:
                    array = array[:, None]
                labels = [str(i + 1) for i in range(array.shape[0])]
                columns = [str(j + 1) for j in range(array.shape[1])]
                matrix_rows.extend(_matrix_rows(key, array, labels, columns, (index, branch, row["iteration"])))
    _table(book, "Attempts", ("attempt", "branch", "status", "accepted", "initial_point_source",
                               "initial_x_S", "initial_x_B", "initial_x_G", "initial_x_C", "reason", "detail"), attempt_rows)
    _table(book, "Iterations", ("attempt", "branch", "iteration", *[f"x_{s}" for s in GROUP_SYMBOLS],
                                 "mu_S", "mu_B", "mu_G", "lambda", "volatility", "objective", "non_bond_exposure",
                                 *[f"risk_share_{s}" for s in GROUP_SYMBOLS], *[f"budget_residual_{s}" for s in GROUP_SYMBOLS],
                                 *[f"stationarity_{s}" for s in GROUP_SYMBOLS], "kkt_residual_max", "step_size", "backtracks",
                                 "linear_solve_residual_max", "line_search_evaluations", "update_accepted", "termination_reason"), iteration_rows)
    if solver_options["trace_level"] == "full":
        _table(book, "LineSearch", ("attempt", "branch", "iteration", "trial", "step_size", "accepted", "reason",
                                     "residual_norm", "required_residual_norm_max", "candidate_x_S", "candidate_x_B", "candidate_x_G", "candidate_x_C",
                                     "candidate_mu_S", "candidate_mu_B", "candidate_mu_G", "candidate_lambda", "detail"), trial_rows)
        _table(book, "KKTMatrices", ("attempt", "branch", "iteration", "matrix", "row", "column", "value"), matrix_rows)
    input_rows = [("model", "min (sqrt(x.T Q x)-target)^2; exact budgets; h.T x<=1; x>0"),
                  ("runner_version", RUNNER_VERSION), ("solver_version", solution.get("model_version", sources.get("solver_version"))),
                  ("group_order", ", ".join(GROUP_ORDER)), ("lookback_days", record["lookback"]),
                  ("rebalance_frequency", record["frequency"]), ("schedule_source_row", record["source_row"]),
                  ("target_volatility", target), ("run_date", datetime.now(timezone(timedelta(hours=8))).date()),
                  *[(f"budget_{g}", b[j]) for j, g in enumerate(GROUP_ORDER)],
                  *[(f"h_{g}", h[j]) for j, g in enumerate(GROUP_ORDER)],
                  *[(key, value) for key, value in solver_options.items()],
                  *[(key, value) for key, value in sources.items()],
                  ("source_Q_policy", "Use original labelled Excel Q; read-only price reconstruction is a consistency check"),
                  ("instrument_scope", "Only the instruments in this period's saved Allocations; no synthetic future instruments"),
                  ("matrix_numbering", "KKT unknowns: x_S,x_B,x_G,x_C,mu_S,mu_B,mu_G[,lambda]; rows: stationarity then constraints")]
    if risk:
        input_rows.extend((key, value) for key, value in risk["fields"].items())
        input_rows.extend((f"source_{key}", value) for key, value in risk["metadata"].items())
        for key in ("source_sha256", "window_sha256", "source_price_hash_matches", "source_schedule_hash_matches"):
            input_rows.append((key, risk.get(key)))
    _table(book, "Inputs", ("field", "value"), input_rows)
    if risk:
        matrices = _matrix_rows("Q_from_source_Excel", risk["Q"], list(GROUP_ORDER), list(GROUP_ORDER))
        if "Sigma" in risk:
            codes = risk["instrument_codes"]
            matrices += _matrix_rows("Sigma_from_prices", risk["Sigma"], codes, codes)
            matrices += _matrix_rows("A", risk["A"], codes, list(GROUP_ORDER))
        _table(book, "RiskMatrices", ("matrix", "row", "column", "value"), matrices)
    if risk and "Sigma" in risk:
        _table(book, "InstrumentReturns", ("return_date", *codes),
               [(day, *row) for day, row in zip(risk["return_dates"], risk["instrument_returns"])])
    _table(book, "Checks", ("check", "error", "tolerance", "passed", "detail"),
           [(r["check"], r.get("error"), r.get("tolerance"), r["passed"], r.get("detail", "")) for r in result.get("checks", [])])
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".optimization-", suffix=".xlsx", dir=path.parent)
    os.close(fd)
    try:
        book.save(temporary)
        if overwrite:
            if path.is_symlink():
                raise ValueError("Refusing to replace an output symlink")
            os.replace(temporary, path)
        else:
            os.link(temporary, path)  # 原子发布且拒绝覆盖已存在的文件。
    finally:
        book.close()
        Path(temporary).unlink(missing_ok=True)


def _target_folder(target: float) -> str:
    percent = target * 100
    label = f"{int(percent):02d}" if percent.is_integer() else format(percent, ".10g").replace(".", "p")
    return f"TargetVol_{label}pct"


def run_history(*, project_root: Path = PROJECT_ROOT, lookback_period: str = LOOKBACK_PERIOD,
                rebalance_frequency: str = REBALANCE_FREQUENCY,
                start_signal_date: Any = START_SIGNAL_DATE, end_signal_date: Any = END_SIGNAL_DATE,
                targets=TARGET_VOLATILITIES, budgets=None, risk_root=None, output_root=None,
                prices_path=None, schedule_path=None, overwrite: bool = OVERWRITE_OUTPUTS,
                failure_policy: str = FAILURE_POLICY, solver_options: Optional[dict] = None) -> dict:
    """独立运行一套窗口/频率。只读输入，每期每目标输出一个 Excel。

    若默认 stop 模式失败，先保存失败工作簿再抛 BatchFailure。
    continue 模式保留所有失败并返回 partial_failure，不把失败期用于回测。
    """
    root = Path(project_root).resolve()
    if lookback_period not in LOOKBACK_WINDOWS or rebalance_frequency not in FREQUENCY_NAMES:
        raise ValueError("Unknown lookback/frequency")
    if failure_policy not in {"stop", "continue"} or not isinstance(overwrite, bool):
        raise ValueError("Invalid failure policy or overwrite setting")
    start = _date(start_signal_date, "start_signal_date", True)
    end = _date(end_signal_date, "end_signal_date", True)
    if start is not None and end is not None and start > end:
        raise ValueError("Start date is after end date")
    targets = tuple(_number(t, "target_volatility") for t in targets)
    if not targets or any(t <= 0 for t in targets) or len(set(targets)) != len(targets):
        raise ValueError("Targets must be distinct, finite, positive numbers")
    if len({_target_folder(t) for t in targets}) != len(targets):
        raise ValueError("Targets would map to duplicate output folders")
    budgets = dict(RISK_BUDGETS if budgets is None else budgets)
    if set(budgets) != set(GROUP_ORDER):
        raise ValueError("Budget labels must be the four agreed groups")
    b = np.array([_number(budgets[g], f"budget_{g}") for g in GROUP_ORDER])
    if np.any(b <= 0) or abs(b.sum() - 1) > 1e-14:
        raise ValueError("Budgets must be positive and sum to one; no renormalization")
    if set(NON_BOND_GROUPS) != {"equity", "gold", "other_commodities"}:
        raise ValueError("The agreed cap is on all three non-bond groups")
    h = np.array([float(g in NON_BOND_GROUPS) for g in GROUP_ORDER])
    options = {"max_iter": MAX_ITERATIONS, "max_backtracks": MAX_BACKTRACKS,
               "residual_tol": KKT_TOLERANCE, "budget_tol": BUDGET_TOLERANCE,
               "cap_tol": CAP_TOLERANCE, "vol_tol": VOLATILITY_TOLERANCE,
               "initial_points": INITIAL_POINTS, "trace_level": TRACE_LEVEL}
    if solver_options:
        unknown = set(solver_options) - (set(options) | {"x0"})
        if unknown:
            raise ValueError(f"Unknown solver settings: {sorted(unknown)}")
        options.update(solver_options)
    if options["trace_level"] not in {"standard", "full"}:
        raise ValueError("trace_level must be standard or full")
    risk_root = Path(risk_root) if risk_root is not None else root / "outputs/risk_estimate"
    output_root = Path(output_root) if output_root is not None else root / "outputs/optimization"
    prices_path = Path(prices_path) if prices_path is not None else root / "data/futures/close_price.csv"
    schedule_path = Path(schedule_path) if schedule_path is not None else root / "data/prep/期货回测安排.xlsx"
    output_root = output_root.resolve()
    for protected in (root / "scripts", root / "data", risk_root, prices_path.parent, schedule_path.parent):
        protected = protected.resolve()
        if output_root == protected or protected in output_root.parents or output_root in protected.parents:
            raise ValueError(f"Output directory must not overlap inputs/code: {protected}")
    risk_module_path = root / "scripts/2.组合优化/1.风险估计.py"
    solver_path = Path(__file__).with_name("3.优化求解.py")
    risk_module = _load_module(risk_module_path, "_allweather_existing_risk_readonly")
    solver = _load_module(solver_path, "_allweather_new_objective_solver")
    prices_hash, schedule_hash = _digest(prices_path), _digest(schedule_path)
    prices = risk_module.load_prices(prices_path)
    schedule = risk_module.load_schedule(schedule_path)
    if _digest(prices_path) != prices_hash or _digest(schedule_path) != schedule_hash:
        raise ValueError("Input files changed while being loaded")
    selected = sorted([r for r in schedule if r["lookback"] == LOOKBACK_WINDOWS[lookback_period]
                       and r["frequency"] == rebalance_frequency
                       and (start is None or r["signal_date"] >= start)
                       and (end is None or r["signal_date"] <= end)], key=lambda r: r["signal_date"])
    if not selected:
        raise ValueError("No preparation schedule rows match the selected settings")
    folder = f"Lookback_{PERIOD_NAMES[lookback_period]}_Rebalance_{FREQUENCY_NAMES[rebalance_frequency]}"
    # 运行前检查全部目的文件，防止先写了一部分才发现旧结果不能覆盖。
    for record in selected:
        for target in targets:
            destination = output_root / folder / _target_folder(target) / f"{record['rebalance_no']:03d}.xlsx"
            if destination.is_symlink() or (destination.exists() and (not overwrite or not destination.is_file())):
                raise FileExistsError(f"Output already exists or is not a regular file: {destination}")
            for parent in destination.parents:
                if parent == output_root.parent:
                    break
                if parent.is_symlink():
                    raise ValueError(f"Output path contains a symlink: {parent}")
    base_sources = {"price_source": prices_path.resolve(), "schedule_source": schedule_path.resolve(),
                    "price_sha256": prices_hash, "schedule_sha256": schedule_hash,
                    "runner_sha256": _digest(Path(__file__)), "solver_sha256": _digest(solver_path),
                    "risk_reader_sha256": _digest(risk_module_path),
                    "solver_version": getattr(solver, "MODEL_VERSION", "new_objective_kkt"),
                    "failure_policy": failure_policy}
    outcomes = []
    for record in selected:
        source = risk_root / folder / f"{record['rebalance_no']:03d}.xlsx"
        sources = {**base_sources, "risk_workbook_source": source.resolve()}
        risk, preparation_error = None, None
        try:
            sources["risk_workbook_sha256"] = _digest(source)
            risk = read_risk_workbook(source)
            risk = prepare_period(risk, record, prices, prices_hash, schedule_hash)
        except Exception as exc:
            # 每期坏 ZIP/XML 等输入异常也必须留痕；不捕获用户中断。
            preparation_error = f"{type(exc).__name__}: {exc}"
        for target in targets:
            destination = output_root / folder / _target_folder(target) / f"{record['rebalance_no']:03d}.xlsx"
            if preparation_error is not None:
                result = {"status": "input_invalid", "error": preparation_error, "solution": {},
                          "groups": [], "instruments": [], "checks": [{"check": "source_input_validation", "passed": False, "detail": preparation_error}]}
            else:
                try:
                    solution = solver.solve_allocation(risk["Q"], b, target, h=h, **options)
                    result = complete_result(risk, solution, b, h, target, options)
                except Exception as exc:
                    result = {"status": "numerical_failure", "error": f"{type(exc).__name__}: {exc}", "solution": {}, "groups": [], "instruments": [],
                              "checks": [{"check": "solver_exception", "passed": False, "detail": str(exc)}]}
            save_workbook(destination, record, risk, result, b, h, target, options, sources, overwrite)
            entry = {"rebalance_no": record["rebalance_no"], "signal_date": record["signal_date"],
                     "target": target, "status": result["status"], "path": destination,
                     "volatility": result.get("volatility")}
            outcomes.append(entry)
            print(f"{record['rebalance_no']:03d} {record['signal_date']:%Y-%m-%d} target={target:.2%} {result['status']} -> {destination.name}")
            if result["status"] not in SUCCESS_STATUSES and failure_policy == "stop":
                raise BatchFailure(f"Stopped after {result['status']}. Diagnostic workbook: {destination}. {result.get('error', result.get('solution', {}).get('reason', ''))}")
    failed = sum(r["status"] not in SUCCESS_STATUSES for r in outcomes)
    return {"status": "partial_failure" if failed else "completed", "failed": failed,
            "periods": len(selected), "outcomes": outcomes, "output_directory": output_root / folder}


def _main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Read existing period risk Excel files and solve the new exact-risk-budget objective.")
    parser.add_argument("--project-root", type=Path, default=PROJECT_ROOT)
    parser.add_argument("--lookback-period", choices=tuple(LOOKBACK_WINDOWS), default=LOOKBACK_PERIOD)
    parser.add_argument("--frequency", choices=tuple(FREQUENCY_NAMES), default=REBALANCE_FREQUENCY)
    parser.add_argument("--start-signal-date", default=START_SIGNAL_DATE)
    parser.add_argument("--end-signal-date", default=END_SIGNAL_DATE)
    parser.add_argument("--targets", nargs="+", type=float, default=TARGET_VOLATILITIES)
    parser.add_argument("--risk-root", type=Path)
    parser.add_argument("--output-root", type=Path)
    parser.add_argument("--prices", type=Path)
    parser.add_argument("--schedule", type=Path)
    parser.add_argument("--overwrite", action="store_true", default=OVERWRITE_OUTPUTS)
    parser.add_argument("--failure-policy", choices=("stop", "continue"), default=FAILURE_POLICY)
    parser.add_argument("--trace-level", choices=("standard", "full"), default=TRACE_LEVEL)
    args = parser.parse_args(argv)
    try:
        result = run_history(project_root=args.project_root, lookback_period=args.lookback_period,
                             rebalance_frequency=args.frequency, start_signal_date=args.start_signal_date,
                             end_signal_date=args.end_signal_date, targets=args.targets, risk_root=args.risk_root,
                             output_root=args.output_root, prices_path=args.prices, schedule_path=args.schedule,
                             overwrite=args.overwrite, failure_policy=args.failure_policy,
                             solver_options={"trace_level": args.trace_level})
    except (OSError, ValueError, TypeError, ImportError, BatchFailure) as exc:
        parser.exit(1, f"Optimization run failed: {exc}\n")
    print(f"Periods: {result['periods']}; optimization results: {len(result['outcomes'])}; failed: {result['failed']}")
    print(f"Output directory: {result['output_directory']}")
    return 1 if result["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(_main())
