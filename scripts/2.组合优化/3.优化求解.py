"""在原四维仓位 x 上求解期货风险预算配置，并返回完整计算轨迹。

问题：min (sqrt(x.T @ Q @ x) - target)**2，x > 0；
      四组风险份额等于 b，最低保证金资金占用 h.T @ x <= 1。
组序为股票、国债、黄金、其他商品，h=(1, bond_margin_floor, 1, 1)。
最低保证金率必须显式传入；求解后由 m_B=(1-x_S-x_G-x_C)/x_B 恢复。

这里只求解，不读取行情、准备表或 Excel，也不保存任何文件。Q 必须
对称正定；输入不合格会明确报错，不自动对称化、改预算或修补协方差。
两个分支分别解「上限乘子为零」和「上限绑定」的原变量 KKT 方程。
三条独立预算等式进入线性系统，最后仍验算全部四组风险份额。

Newton 法配合正域回溯，下降的是 KKT 残差范数。它不保证从任意初值
收敛；数值失败不等于模型无解。通过验算的解具有最优性的数学依据，
与 Newton 算法能否找到该解是两件事。本模块不使用对数目标、归一化
变量或按比例缩放来替代原四变量问题的求解。

公开接口：solve_allocation(...)。evaluate_kkt(...) 可用于核对导数。
返回值中的数组均为 NumPy 数组；由调用方负责按需要保存结果。
Python 3.9+，仅依赖 NumPy。导入不产生文件，直接运行仅打印用法。
"""

from numbers import Integral, Real

import numpy as np


MODEL_VERSION = "original_x_bond_margin_floor_1.0"


def _real_scalar(value, name, positive=True):
    """拒绝 bool、字符串、NaN/Inf；不把 True 偷偷当作数值 1。"""
    if isinstance(value, (bool, np.bool_)) or not isinstance(value, Real):
        raise ValueError("%s 必须是实数，不能是 bool 或字符串。" % name)
    try:
        result = float(value)
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError("%s 不能转换为有限实数。" % name) from exc
    if not np.isfinite(result) or (positive and result <= 0):
        raise ValueError("%s 必须是有限正数。" % name)
    return result


def _integer(value, name, minimum):
    if isinstance(value, (bool, np.bool_)) or not isinstance(value, Integral):
        raise ValueError("%s 必须是整数，不能是 bool。" % name)
    if value < minimum:
        raise ValueError("%s 必须不小于 %d。" % (name, minimum))
    return int(value)


def _real_array(value, name, shape):
    """先逐项检查原类型，防止混合列表中的 bool 在转换时被掩盖。"""
    try:
        raw = np.asarray(value, dtype=object)
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError("%s 不是有效的实数数组。" % name) from exc
    if raw.shape != shape:
        raise ValueError("%s 的形状必须为 %s。" % (name, shape))
    if any(isinstance(item, (bool, np.bool_)) or not isinstance(item, Real)
           for item in raw.flat):
        raise ValueError("%s 只能包含实数，不能包含 bool、复数或字符串。" % name)
    try:
        array = np.asarray(value, dtype=float).copy()
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError("%s 不能转换为实数数组。" % name) from exc
    if not np.isfinite(array).all():
        raise ValueError("%s 不能包含 NaN 或 Inf。" % name)
    return array


def _inputs(Q, b, target, bond_margin_floor):
    Q = _real_array(Q, "Q", (4, 4))
    b = _real_array(b, "b", (4,))
    margin_floor = _real_scalar(bond_margin_floor, "bond_margin_floor")
    if not np.isfinite(1.0 / margin_floor):
        raise ValueError("bond_margin_floor 的倒数必须是有限的国债杠杆上限。")
    h = np.array([1.0, margin_floor, 1.0, 1.0])
    # 与已验证原型一致：精确检查对称，不悄悄用 (Q+Q.T)/2 替换输入。
    if not np.array_equal(Q, Q.T):
        raise ValueError("Q 必须对称；本模块不会自动对称化，请检查协方差构造。")
    try:
        np.linalg.cholesky(Q)
    except np.linalg.LinAlgError as exc:
        raise ValueError("Q 必须正定；本求解器不支持奇异或非正定 Q。") from exc
    if np.any(b <= 0) or not np.isclose(b.sum(), 1.0, rtol=0, atol=1e-14):
        raise ValueError("b 的四项必须严格为正且合计为 1；不自动归一化预算。")
    return Q, b, _real_scalar(target, "target"), h


def _initial_points(Q, b, target, x0, initial_points):
    """显式初值按用户顺序尝试；未给初值时才使用三个公开默认猜测。"""
    supplied = x0 is not None or initial_points is not None
    points = []
    if x0 is not None:
        points.append((x0, "x0"))
    if initial_points is not None:
        if isinstance(initial_points, (str, bytes, dict)):
            raise ValueError("initial_points 必须是四维正初值组成的可迭代对象。")
        try:
            points.extend((point, "initial_points[%d]" % index)
                          for index, point in enumerate(initial_points))
        except TypeError as exc:
            raise ValueError("initial_points 必须是初值的可迭代对象。") from exc
    if not supplied:
        # 第二项仅作对角近似的起点，不将它认定为解，也不缩放计算结果。
        with np.errstate(over="raise", divide="raise", invalid="raise"):
            try:
                diagonal_guess = target * np.sqrt(b / np.diag(Q))
            except FloatingPointError as exc:
                raise ValueError("默认初值计算超出浮点范围；请提供正的 x0。") from exc
        points = [
            (np.array([0.30, 1.50, 0.25, 0.20]), "default_fixed"),
            (diagonal_guess, "default_diagonal_guess"),
            (np.full(4, 0.25), "default_equal_guess"),
        ]
    if not points:
        raise ValueError("至少提供一个严格为正的四维初值。")
    checked = []
    for point, source in points:
        point = _real_array(point, source, (4,))
        if np.any(point <= 0):
            raise ValueError("%s 的四个初始仓位必须严格为正。" % source)
        if not any(np.array_equal(point, other[0]) for other in checked):
            checked.append((point, source))
    return checked


def evaluate_kkt(z, Q, b, target, h, active_cap):
    """计算 7/8 维 KKT 残差和解析 Jacobian，不改变任何输入。

    z[:4] 是原仓位 x，z[4:7] 是前三条预算等式的乘子；绑定分支
    的 z[7] 是上限乘子。前三组依次为 S、B、G，第四组 C 留作验算。
    本函数供求解器内部反复调用，Q/b/target/h 已由入口严格检查。
    """
    z = np.asarray(z, dtype=float)
    if z.shape != (8 if active_cap else 7,) or not np.isfinite(z).all():
        raise ValueError("KKT 向量的尺寸或有限性不正确。")
    x, multipliers = z[:4], z[4:]
    if np.any(x <= 0):
        raise ValueError("KKT 计算要求 x 的每个分量大于零。")
    with np.errstate(over="raise", divide="raise", invalid="raise"):
        qx = Q @ x
        variance = float(x @ qx)
        if not np.isfinite(variance) or variance <= 0:
            raise FloatingPointError("组合方差不是有限正数。")
        volatility = float(np.sqrt(variance))
        objective = float((volatility - target) ** 2)
        gradient = 2 * (1 - target / volatility) * qx
        objective_hessian = (
            2 * (1 - target / volatility) * Q
            + 2 * target / volatility**3 * np.outer(qx, qx)
        )
        eye = np.eye(4)
        risk_numerators = x * qx
        risk_equations = risk_numerators - b * variance
        risk_shares = risk_numerators / variance
        budget_jacobian = np.array([
            eye[g] * qx[g] + x[g] * Q[g] - 2 * b[g] * qx
            for g in range(3)
        ])
        budget_hessians = np.array([
            np.outer(eye[g], Q[g]) + np.outer(Q[g], eye[g]) - 2 * b[g] * Q
            for g in range(3)
        ])
        lagrangian_hessian = objective_hessian + np.einsum(
            "g,gij->ij", multipliers[:3], budget_hessians
        )
        jacobian = budget_jacobian
        constraints = risk_equations[:3]
        cap_exposure = float(h @ x)
        if active_cap:
            jacobian = np.vstack([jacobian, h])
            constraints = np.r_[constraints, cap_exposure - 1]
        stationarity = gradient + jacobian.T @ multipliers
        residual = np.r_[stationarity, constraints]
        kkt_matrix = np.block([
            [lagrangian_hessian, jacobian.T],
            [jacobian, np.zeros((len(constraints), len(constraints)))],
        ])
    state = {
        "qx": qx, "variance": variance, "volatility": volatility,
        "objective": objective, "objective_gradient": gradient,
        "objective_hessian": objective_hessian,
        "risk_numerators": risk_numerators, "risk_shares": risk_shares,
        "all_risk_equations": risk_equations,
        "constraint_jacobian": jacobian,
        "lagrangian_hessian": lagrangian_hessian,
        "stationarity": stationarity, "residual": residual,
        "kkt_matrix": kkt_matrix, "capital_used_at_floor": cap_exposure,
    }
    if any(not np.isfinite(value).all() for value in state.values()):
        raise FloatingPointError("KKT 状态或矩阵包含非有限数值。")
    return state


def _history_record(iteration, z, state, trace_level):
    """每条记录表示本轮更新前状态；步长/方向描述从本轮走向下一轮。"""
    record = {
        "iteration": iteration, "x": z[:4].copy(),
        "multipliers": z[4:].copy(),
        "volatility": state["volatility"], "objective": state["objective"],
        "variance": state["variance"], "qx": state["qx"].copy(),
        "risk_shares": state["risk_shares"].copy(),
        "all_risk_equations": state["all_risk_equations"].copy(),
        "stationarity": state["stationarity"].copy(),
        "residual_max": float(np.max(np.abs(state["residual"]))),
        "capital_used_at_floor": state["capital_used_at_floor"],
        "step_size": None, "backtracks": 0, "line_search_evaluations": 0,
        "direction": None, "linear_solve_residual_max": None,
        "update_accepted": False, "termination_reason": None,
    }
    if trace_level == "full":
        for name in ("objective_gradient", "objective_hessian", "constraint_jacobian",
                     "lagrangian_hessian", "kkt_matrix", "residual"):
            record[name] = state[name].copy()
        record["rhs"] = -state["residual"].copy()
        record["line_search_trials"] = []
    return record


def _newton(Q, b, target, h, initial_x, active_cap, max_iter,
            residual_tol, max_backtracks, trace_level):
    """解原 KKT 方程；回溯只保证正域与残差下降，不承诺全局收敛。"""
    z = np.r_[initial_x, np.zeros(4 if active_cap else 3)]
    initial_z = z.copy()
    history = []

    def failure(reason, detail):
        if history and history[-1]["termination_reason"] is None:
            history[-1]["termination_reason"] = reason
        return {
            "status": "numerical_failure", "reason": reason, "detail": detail,
            "active_cap": active_cap, "initial_x": initial_x.copy(),
            "initial_z": initial_z.copy(),
            "initial_multipliers": initial_z[4:].copy(),
            "last_z": z.copy(), "history": history,
        }

    for iteration in range(max_iter + 1):
        try:
            state = evaluate_kkt(z, Q, b, target, h, active_cap)
        except (ValueError, FloatingPointError, OverflowError, ZeroDivisionError) as exc:
            return failure("evaluation_failed", str(exc))
        residual = state["residual"]
        record = _history_record(iteration, z, state, trace_level)
        history.append(record)
        if record["residual_max"] <= residual_tol:
            record["termination_reason"] = "kkt_residual_tolerance"
            return {
                "status": "kkt_root", "active_cap": active_cap,
                "initial_x": initial_x.copy(), "initial_z": initial_z.copy(),
                "initial_multipliers": initial_z[4:].copy(),
                "z": z.copy(), "history": history,
            }
        if iteration == max_iter:
            return failure("iteration_limit", "已达到 %d 次 Newton 更新上限。" % max_iter)
        try:
            direction = np.linalg.solve(state["kkt_matrix"], -residual)
        except np.linalg.LinAlgError as exc:
            return failure("singular_kkt_system", str(exc))
        if not np.isfinite(direction).all():
            return failure("nonfinite_direction", "Newton 方向包含非有限数值。")
        record["direction"] = direction.copy()
        with np.errstate(over="ignore", invalid="ignore"):
            linear_error = state["kkt_matrix"] @ direction + residual
            residual_norm = float(np.linalg.norm(residual))
        if not np.isfinite(linear_error).all():
            return failure("nonfinite_linear_residual", "线性方程残差非有限。")
        record["linear_solve_residual_max"] = float(np.max(np.abs(linear_error)))
        if not np.isfinite(residual_norm):
            return failure("nonfinite_residual_norm", "KKT 残差范数超出浮点范围。")
        record["residual_norm"] = residual_norm
        step = 1.0
        accepted = False
        for trial_index in range(max_backtracks):
            with np.errstate(over="ignore", invalid="ignore"):
                candidate = z + step * direction
            threshold = (1 - 1e-4 * step) * residual_norm
            trial = {
                "trial_no": trial_index + 1, "step_size": step,
                "accepted": False, "reason": None,
                "residual_norm": None, "required_residual_norm_max": threshold,
            }
            if trace_level == "full":
                trial["candidate_x"] = candidate[:4].copy()
                trial["candidate_multipliers"] = candidate[4:].copy()
            if not np.isfinite(candidate).all():
                trial["reason"] = "nonfinite_candidate"
            elif np.any(candidate[:4] <= 0):
                trial["reason"] = "nonpositive_exposure"
            else:
                try:
                    next_state = evaluate_kkt(candidate, Q, b, target, h, active_cap)
                    with np.errstate(over="ignore", invalid="ignore"):
                        new_norm = float(np.linalg.norm(next_state["residual"]))
                    trial["residual_norm"] = new_norm
                    trial["residual_max"] = float(np.max(np.abs(next_state["residual"])))
                    if not np.isfinite(new_norm):
                        trial["reason"] = "nonfinite_residual_norm"
                    elif new_norm <= threshold:
                        trial["accepted"] = True
                        trial["reason"] = "sufficient_residual_decrease"
                        accepted = True
                    else:
                        trial["reason"] = "insufficient_residual_decrease"
                except (ValueError, FloatingPointError, OverflowError, ZeroDivisionError) as exc:
                    trial["reason"] = "candidate_evaluation_failed"
                    trial["detail"] = str(exc)
            record["line_search_evaluations"] = trial_index + 1
            if trace_level == "full":
                record["line_search_trials"].append(trial)
            if accepted:
                record["step_size"] = step
                record["backtracks"] = trial_index
                record["update_accepted"] = True
                break
            record["backtracks"] = trial_index + 1
            step *= 0.5
        if not accepted:
            return failure("no_residual_descent", "正域回溯耗尽，未找到足够下降的步长。")
        z = candidate
    raise AssertionError("不应到达 Newton 循环之外。")


def _postcheck(root, Q, b, target, h, residual_tol, budget_tol, cap_tol, vol_tol):
    """残差小不等于合格解：同时验算第四预算、上限、乘子和分支条件。"""
    z = root["z"]
    active_cap = root["active_cap"]
    state = evaluate_kkt(z, Q, b, target, h, active_cap)
    x = z[:4]
    cap = state["capital_used_at_floor"]
    multiplier = float(z[-1]) if active_cap else 0.0
    errors = {
        "stationarity_max": float(np.max(np.abs(state["stationarity"]))),
        "kkt_residual_max": float(np.max(np.abs(state["residual"]))),
        "all_four_budget_share_error_max": float(np.max(np.abs(state["risk_shares"] - b))),
        "all_four_budget_equation_error_max": float(np.max(np.abs(state["all_risk_equations"]))),
        "cap_violation": max(0.0, cap - 1),
        "cap_equality_error": abs(cap - 1) if active_cap else None,
        "cap_multiplier_negative_part": max(0.0, -multiplier),
        "complementarity_error": abs(multiplier * (cap - 1)),
        "radial_stationarity_error": abs(
            2 * state["volatility"] * (state["volatility"] - target) + multiplier * cap
        ),
        "target_error": abs(state["volatility"] - target),
    }
    failed = []
    # 极端数值下也不能让 NaN/Inf 因比较结果为 False 而绕过验收。
    with np.errstate(over="ignore", divide="ignore", invalid="ignore"):
        gross_exposure = float(x.sum())
        risk_contributions = state["risk_numerators"] / state["volatility"]
    if (any(value is not None and not np.isfinite(value) for value in errors.values())
            or not np.isfinite(gross_exposure)
            or not np.isfinite(risk_contributions).all()):
        failed.append("nonfinite_postcheck")
    if np.any(x <= 0):
        failed.append("positive_domain")
    if errors["kkt_residual_max"] > residual_tol:
        failed.append("kkt_residual")
    if errors["all_four_budget_share_error_max"] > budget_tol:
        failed.append("all_four_risk_budgets")
    # 无论哪个分支，上限失败都不能令 passed=True。
    if errors["cap_violation"] > cap_tol:
        failed.append("cap_violation")
    if active_cap:
        if errors["cap_equality_error"] > cap_tol:
            failed.append("active_cap_equality")
        if multiplier < -residual_tol:
            failed.append("negative_cap_multiplier")
        if errors["complementarity_error"] > cap_tol:
            failed.append("complementarity")
        if state["volatility"] > target + vol_tol:
            failed.append("active_point_above_target")
    elif errors["target_error"] > vol_tol:
        failed.append("inactive_point_not_at_target")
    non_bond = float(x[[0, 2, 3]].sum())
    bond_capital = 1.0 - non_bond
    margin_rate = bond_capital / float(x[1])
    leverage = float(x[1]) / bond_capital if bond_capital > 0 else None
    # 超限的未绑定根仍可作为绑定分支初值；只有可行候选才验算恢复的 m_B。
    if errors["cap_violation"] <= cap_tol:
        if (not np.isfinite(margin_rate) or margin_rate <= 0 or leverage is None
                or not np.isfinite(leverage)):
            failed.append("invalid_bond_capital")
        elif max(0.0, (h[1] - margin_rate) * x[1]) > cap_tol:
            failed.append("bond_margin_floor")
    return {
        "passed": not failed, "failed_checks": failed, "errors": errors,
        "x": x.copy(), "budget_multipliers": z[4:7].copy(),
        "cap_multiplier": multiplier, "active_cap": active_cap,
        "capital_used_at_floor": cap, "gross_exposure": gross_exposure,
        "volatility": state["volatility"], "variance": state["variance"],
        "objective": state["objective"], "qx": state["qx"].copy(),
        "risk_shares": state["risk_shares"].copy(),
        "risk_numerators": state["risk_numerators"].copy(),
        "risk_contributions": risk_contributions,
        "bond_margin_floor": float(h[1]), "bond_leverage_cap": 1.0 / float(h[1]),
        "bond_margin_rate": margin_rate, "bond_leverage": leverage,
        "non_bond_exposure": non_bond, "bond_exposure": float(x[1]),
        "bond_capital_weight": bond_capital, "capital_usage": non_bond + margin_rate * x[1],
    }


def solve_allocation(Q, b, target, *, bond_margin_floor, x0=None, initial_points=None,
                     max_iter=100, residual_tol=1e-13, budget_tol=1e-8,
                     cap_tol=1e-10, vol_tol=1e-10, max_backtracks=60,
                     trace_level="full"):
    """返回 target_achieved / cap_limited / numerical_failure。

    Q、b 和 x 按 S,B,G,C 排列；h 由显式传入的 bond_margin_floor 构造。
    max_iter 是每次尝试的最大更新次数；
    max_backtracks 是每轮最多试探步长数（包含第一次完整步长）。四个
    tolerance 都必须是有限正数。full 记录每轮全部矩阵和试探步原因；
    standard 省略大矩阵与逐次试探表，仍保留初值、乘子和逐轮关键状态。

    显式给出 x0/initial_points 后，不擅自追加默认猜测。未绑定分支的
    合格方程根若仅因超限被拒绝，可原样作为绑定分支的额外起点：只用
    作初值，不做缩放。attempts 记录每次分支尝试及接受或拒绝理由。
    返回的「通过」表示在给定容差内通过验算，不声称算法全局收敛。
    """
    Q, b, target, h = _inputs(Q, b, target, bond_margin_floor)
    max_iter = _integer(max_iter, "max_iter", 0)
    max_backtracks = _integer(max_backtracks, "max_backtracks", 1)
    residual_tol = _real_scalar(residual_tol, "residual_tol")
    budget_tol = _real_scalar(budget_tol, "budget_tol")
    cap_tol = _real_scalar(cap_tol, "cap_tol")
    vol_tol = _real_scalar(vol_tol, "vol_tol")
    if not isinstance(trace_level, str) or trace_level not in ("full", "standard"):
        raise ValueError("trace_level 只能为 full 或 standard。")
    points = _initial_points(Q, b, target, x0, initial_points)
    attempts = []
    inactive_candidate = None

    def attempt(point, source, active_cap):
        root = _newton(Q, b, target, h, point, active_cap, max_iter,
                       residual_tol, max_backtracks, trace_level)
        root.update({
            "attempt_no": len(attempts) + 1,
            "branch": "active_cap" if active_cap else "inactive_cap",
            "initial_point_source": source,
            "accepted": False, "acceptance_reason": None, "rejection_reason": None,
        })
        attempts.append(root)
        if root["status"] != "kkt_root":
            root["rejection_reason"] = root["reason"]
            return root, None
        try:
            check = _postcheck(root, Q, b, target, h, residual_tol,
                               budget_tol, cap_tol, vol_tol)
        except (ValueError, FloatingPointError, OverflowError, ZeroDivisionError) as exc:
            root.update(status="numerical_failure", reason="postcheck_evaluation_failed",
                        detail=str(exc), rejection_reason="postcheck_evaluation_failed")
            return root, None
        root["postcheck"] = check
        if not check["passed"]:
            root.update(status="rejected_candidate", reason="postcheck_failed",
                        detail=", ".join(check["failed_checks"]),
                        rejection_reason=", ".join(check["failed_checks"]))
        return root, check

    def success(root, check, status):
        if not check["passed"] or check["capital_used_at_floor"] > 1 + cap_tol:
            raise AssertionError("禁止将未通过验算或超限的仓位返回为成功。")
        root["accepted"] = True
        root["acceptance_reason"] = (
            "全部验算通过，目标波动在容差内达到。" if status == "target_achieved"
            else "全部验算通过，上限绑定且波动低于目标。"
        )
        return {
            "status": status, "target": target, **check,
            "model_version": MODEL_VERSION, "trace_level": trace_level,
            "inactive_candidate": inactive_candidate, "attempts": attempts,
            "global_optimality_basis": (
                "Q 正定且预算严格为正时，正的精确预算配置具有唯一比例关系。"
                "验算后的可行精确目标点具有零目标值；或上限绑定且波动低于目标的"
                "合格 KKT 点是该约束集合中最接近目标者。这是解的最优性依据，"
                "并不表示非凸 KKT Newton 法从任意初值全局收敛。"
            ),
        }

    for point, source in points:
        root, check = attempt(point, source, False)
        if check is None:
            continue
        if check["passed"]:
            inactive_candidate = check
            return success(root, check, "target_achieved")
        if check["failed_checks"] == ["cap_violation"]:
            # 方程根本身可靠，但它不是可行组合；转而求上限绑定分支。
            inactive_candidate = check
            break

    active_points = list(points)
    if inactive_candidate is not None:
        point = inactive_candidate["x"]
        if not any(np.array_equal(point, other[0]) for other in active_points):
            active_points.append((point, "inactive_root_without_rescaling"))
    for point, source in active_points:
        root, check = attempt(point, source, True)
        if check is None or not check["passed"]:
            continue
        status = "target_achieved" if abs(check["volatility"] - target) <= vol_tol else "cap_limited"
        return success(root, check, status)
    return {
        "status": "numerical_failure", "target": target,
        "model_version": MODEL_VERSION, "trace_level": trace_level,
        "reason": "No attempted KKT root passed all required checks",
        "message": "所有已尝试的 KKT 候选都未通过验算；这是数值失败，不是模型无解的结论。",
        "inactive_candidate": inactive_candidate, "attempts": attempts,
    }


if __name__ == "__main__":
    print("本程序提供 solve_allocation(Q, b, target, bond_margin_floor=...)；请由 2.逐期运行.py 导入调用。")
    print("只返回仓位、验算及迭代轨迹，不读取行情、不保存文件。")
