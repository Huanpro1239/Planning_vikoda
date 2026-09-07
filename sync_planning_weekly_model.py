"""Adapter SharePoint -> engine ``Ke hoach SX tuan`` -> proposal/report."""
from __future__ import annotations

import calendar
import hashlib
import math
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from openpyxl.utils.datetime import to_excel
from io import BytesIO
from typing import Any

from openpyxl import load_workbook

from weekly_planning_engine import PlannerPolicy, WeeklyCalculatedRow, WeeklyInputRow, build_daily_plan, calculate_rows

PLANNING_SHEET = "Ke_hoach_SX"
MASTER_SHEET = "Danh_muc"
START_COLUMN = 19  # S
EPS = 1e-6
BALANCE_EPS = 1e-5


@dataclass(frozen=True)
class WeeklyAnalysis:
    calculated: list[WeeklyCalculatedRow]
    daily_plan: list[Any]
    policy_warnings: list[dict[str, Any]]
    changed_cells: int
    period_year: int
    period_month: int


def _norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def _num(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, bool):
        raise RuntimeError("TRUE/FALSE không phải số kế hoạch.")
    result = float(value)
    if not math.isfinite(result):
        raise RuntimeError(f"Giá trị không hữu hạn: {value!r}")
    return result


def _code(value: Any) -> int:
    return int(float(value))


def _header_col(ws, names: set[str]) -> int | None:
    wanted = {_norm(name) for name in names}
    for col in range(1, ws.max_column + 1):
        if _norm(ws.cell(1, col).value) in wanted:
            return col
    return None


def _debt_mode(value: Any) -> str | None:
    text = _norm(value).replace("-", "_").replace(" ", "_")
    if not text:
        return None
    if text in {"subtract_book_on_debt", "tru_ton_so", "trừ_tồn_sổ"}:
        return "SUBTRACT_BOOK_ON_DEBT"
    if text in {"ignore_book_on_debt", "khong_tru_ton_so", "không_trừ_tồn_sổ"}:
        return "IGNORE_BOOK_ON_DEBT"
    raise RuntimeError(f"Debt mode không hợp lệ: {value!r}")


def _profile(value: Any) -> str | None:
    text = _norm(value).replace("-", "_").replace(" ", "_")
    if not text:
        return None
    if text in {"continuous", "lien_tuc", "liên_tục"}:
        return "CONTINUOUS"
    if text in {"spread_non_sunday", "rai_khong_chu_nhat", "rải_không_chủ_nhật"}:
        return "SPREAD_NON_SUNDAY"
    raise RuntimeError(f"Schedule profile không hợp lệ: {value!r}")


def _same(current: Any, target: Any) -> bool:
    if target is None:
        return current in (None, "")
    if isinstance(target, datetime):
        if isinstance(current, datetime):
            return abs((current - target).total_seconds()) < 1
        if isinstance(current, date):
            return current == target.date()
        if isinstance(current, (int, float)):
            return math.isclose(float(current), float(to_excel(target)), rel_tol=0.0, abs_tol=1.0/86400.0)
        return False
    if isinstance(current, (int, float)) and isinstance(target, (int, float)):
        return math.isclose(float(current), float(target), rel_tol=1e-10, abs_tol=1e-7)
    return current == target


def _read_inputs(workbook_bytes: bytes, *, plan_year: int, plan_month: int):
    wb = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        for name in (PLANNING_SHEET, MASTER_SHEET):
            if name not in wb.sheetnames:
                raise RuntimeError(f"Không tìm thấy sheet {name!r}.")
        planning, master = wb[PLANNING_SHEET], wb[MASTER_SHEET]
        debt_col = _header_col(master, {"Debt mode", "Cách tính nợ"})
        profile_col = _header_col(master, {"Schedule profile", "Profile lịch"})
        master_data: dict[int, dict[str, Any]] = {}
        for r in range(2, master.max_row + 1):
            raw = master.cell(r, 1).value
            if raw in (None, ""):
                continue
            code = _code(raw)
            if code in master_data:
                raise RuntimeError(f"Mã {code} bị lặp trong {MASTER_SHEET}!A.")
            master_data[code] = {
                "name": str(master.cell(r, 2).value or ""),
                "uom": str(master.cell(r, 3).value or ""),
                "batch": _num(master.cell(r, 4).value),
                "per_shift": _num(master.cell(r, 5).value),
                "line": str(master.cell(r, 6).value or "").strip(),
                "group": str(master.cell(r, 7).value or "").strip(),
                "classification": str(master.cell(r, 8).value or "").strip(),
                "mold": _num(master.cell(r, 9).value),
                "leadtime": _num(master.cell(r, 10).value),
                "debt_mode": _debt_mode(master.cell(r, debt_col).value) if debt_col else None,
                "profile": _profile(master.cell(r, profile_col).value) if profile_col else None,
            }

        rows: list[WeeklyInputRow] = []
        warnings: list[dict[str, Any]] = []
        spread: set[int] = set()
        for r in range(2, planning.max_row + 1):
            raw = planning.cell(r, 1).value
            if raw in (None, ""):
                continue
            code = _code(raw)
            meta = master_data.get(code)
            if meta is None:
                raise RuntimeError(f"Mã {code} không có trong {MASTER_SHEET}.")
            debt = _num(planning.cell(r, 14).value)
            mode = meta["debt_mode"]
            if mode is None:
                mode = "SUBTRACT_BOOK_ON_DEBT"  # proposal-only fallback
                if debt > EPS:
                    warnings.append({"code": str(code), "type": "MISSING_DEBT_MODE", "message": "Thiếu Danh_muc!Debt mode; proposal tạm dùng SUBTRACT_BOOK_ON_DEBT."})
            if meta["line"] == "Galon":
                if meta["profile"] == "SPREAD_NON_SUNDAY":
                    spread.add(code)
                elif meta["profile"] is None:
                    warnings.append({"code": str(code), "type": "MISSING_SCHEDULE_PROFILE", "message": "Thiếu Danh_muc!Schedule profile cho Galon; proposal tạm dùng CONTINUOUS."})
            fc = _num(planning.cell(r, 12).value)
            rows.append(WeeklyInputRow(
                source_row=r, ma_sp=code, ten_sp=meta["name"], don_vi_tinh=meta["uom"],
                sl_me=meta["batch"], sl_ca=meta["per_shift"], chuyen=meta["line"],
                nhom_sp=meta["group"], phan_loai=meta["classification"], quy_cach=meta["mold"],
                shifts_per_day=_num(planning.cell(r, 9).value), ton_dau_thuc_te=_num(planning.cell(r, 10).value),
                ton_dau_so_sach=_num(planning.cell(r, 11).value), fc=fc,
                ton_cuoi_du_kien=_num(planning.cell(r, 13).value), no_kho=debt,
                avg_daily_sales=fc / 26.0 if abs(fc) > EPS else 0.0,
                leadtime=meta["leadtime"], debt_formula_mode=mode,
            ))
        return rows, frozenset(spread), warnings
    finally:
        wb.close()


def analyze_weekly_workbook(workbook_bytes: bytes, *, plan_year: int, plan_month: int) -> WeeklyAnalysis:
    rows, spread, warnings = _read_inputs(workbook_bytes, plan_year=plan_year, plan_month=plan_month)
    calculated = calculate_rows(rows, period_year=plan_year, period_month=plan_month)
    daily = build_daily_plan(calculated, policy=PlannerPolicy(spread_product_codes=spread))
    lookup: dict[tuple[int, date], float] = defaultdict(float)
    for item in daily:
        lookup[(item.ma_sp, item.date)] += float(item.qty)
    wb = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        ws = wb[PLANNING_SHEET]
        changed = 0
        days = calendar.monthrange(plan_year, plan_month)[1]
        for calc in calculated:
            r = calc.input.source_row
            for c, target in ((15, calc.p_need), (16, calc.q_rounded), (17, calc.production_days), (18, calc.start_datetime)):
                if not _same(ws.cell(r, c).value, target):
                    changed += 1
            for d in range(1, days + 1):
                target = lookup.get((calc.input.ma_sp, date(plan_year, plan_month, d)), 0.0) or None
                if not _same(ws.cell(r, START_COLUMN + d - 1).value, target):
                    changed += 1
        return WeeklyAnalysis(calculated, daily, warnings, changed, plan_year, plan_month)
    finally:
        wb.close()


def patch_weekly_workbook(workbook_bytes: bytes, analysis: WeeklyAnalysis) -> bytes:
    wb = load_workbook(BytesIO(workbook_bytes), data_only=False)
    try:
        ws = wb[PLANNING_SHEET]
        lookup: dict[tuple[int, date], float] = defaultdict(float)
        for item in analysis.daily_plan:
            lookup[(item.ma_sp, item.date)] += float(item.qty)
        days = calendar.monthrange(analysis.period_year, analysis.period_month)[1]
        for calc in analysis.calculated:
            r = calc.input.source_row
            ws.cell(r, 15, calc.p_need); ws.cell(r, 16, calc.q_rounded)
            ws.cell(r, 17, calc.production_days); ws.cell(r, 18, calc.start_datetime)
            for d in range(1, days + 1):
                qty = lookup.get((calc.input.ma_sp, date(analysis.period_year, analysis.period_month, d)), 0.0)
                ws.cell(r, START_COLUMN + d - 1).value = qty if qty > EPS else None
        out = BytesIO(); wb.save(out); return out.getvalue()
    finally:
        wb.close()


def _mass_balance(analysis: WeeklyAnalysis):
    scheduled: dict[str, float] = defaultdict(float)
    for item in analysis.daily_plan:
        scheduled[str(item.ma_sp)] += float(item.qty)
    mass, carryovers = {}, []
    for calc in analysis.calculated:
        code = str(calc.input.ma_sp)
        planned = max(0.0, float(calc.q_rounded)); done = scheduled.get(code, 0.0)
        if done > planned + BALANCE_EPS:
            raise RuntimeError(f"Mã {code} scheduled={done} vượt Q={planned}.")
        carry = max(0.0, planned - done)
        if carry <= BALANCE_EPS: carry = 0.0
        if carry: carryovers.append(code)
        mass[code] = {"uom": calc.input.don_vi_tinh, "planned_qty": planned, "scheduled_qty": done, "carryover_qty": carry, "balanced_qty": done + carry}
    return mass, sorted(carryovers)


def _validate_shared_machine(analysis: WeeklyAnalysis, policy: PlannerPolicy | None = None):
    """Kiểm tra capacity vật lý của máy chung KHS + PET 9000.

    Báo cáo cũ đánh dấu resource_validation=True vô điều kiện. Điều đó có thể
    làm proposal xanh dù hai line đã bị xếp song song. Validation này quy đổi
    sản lượng ngày về số ca và kiểm tra tổng usage của *một* máy chung.
    """

    active_policy = policy or PlannerPolicy()
    selected = [
        calc
        for calc in analysis.calculated
        if calc.input.chuyen in active_policy.serialized_lines
        and calc.q_rounded > EPS
    ]
    if not selected:
        return {
            "ok": True,
            "state": "not_used",
            "resource": "KHS/PET 9000 shared machine",
            "capacity_shifts_per_day": 0.0,
            "peak_usage_shifts": 0.0,
            "over_capacity_dates": [],
        }

    capacity = selected[0].input.shifts_per_day
    if capacity <= EPS:
        return {
            "ok": False,
            "state": "invalid_capacity",
            "resource": "KHS/PET 9000 shared machine",
            "capacity_shifts_per_day": capacity,
            "peak_usage_shifts": 0.0,
            "over_capacity_dates": [],
        }

    if any(abs(calc.input.shifts_per_day - capacity) > EPS for calc in selected[1:]):
        return {
            "ok": False,
            "state": "inconsistent_shift_calendar",
            "resource": "KHS/PET 9000 shared machine",
            "capacity_shifts_per_day": capacity,
            "peak_usage_shifts": 0.0,
            "over_capacity_dates": [],
        }

    by_code = {calc.input.ma_sp: calc for calc in selected}
    usage: dict[date, float] = defaultdict(float)
    for item in analysis.daily_plan:
        calc = by_code.get(item.ma_sp)
        if calc is not None:
            usage[item.date] += float(item.qty) / calc.input.sl_ca

    peak = max(usage.values(), default=0.0)
    over = sorted(
        current.isoformat()
        for current, used in usage.items()
        if used > capacity + EPS
    )
    return {
        "ok": not over,
        "state": "shared_machine_passed" if not over else "shared_machine_over_capacity",
        "resource": "KHS/PET 9000 shared machine",
        "capacity_shifts_per_day": capacity,
        "peak_usage_shifts": peak,
        "over_capacity_dates": over,
    }


def build_weekly_schedule_report(workbook_bytes: bytes, analysis: WeeklyAnalysis, *, input_revision=None):
    mass, carryovers = _mass_balance(analysis)
    lines = sorted({calc.input.chuyen for calc in analysis.calculated})
    shared_machine = _validate_shared_machine(analysis)
    shared_name = shared_machine["resource"]
    status = {
        "monthly_quantity": {"ok": not carryovers, "state": "complete" if not carryovers else "carryover", "carryover_skus": carryovers},
        "resource_validation": {
            "ok": shared_machine["ok"],
            "state": shared_machine["state"],
            "validated_resources": [shared_name],
            "failed_resources": [] if shared_machine["ok"] else [shared_name],
            "shared_machine": shared_machine,
        },
        "service": {"ok": not analysis.policy_warnings, "state": "rules_complete" if not analysis.policy_warnings else "policy_metadata_missing", "stockout_skus": []},
    }
    resources = {line: {"meta": {"mode": "ke_hoach_sx_tuan_reference_rules"}} for line in lines}
    resources[shared_name] = {"meta": shared_machine}
    return {
        "schema_version": 4, "algorithm": "ke_hoach_sx_tuan_v1",
        "plan_month": f"{analysis.period_year:04d}-{analysis.period_month:02d}",
        "input_revision": dict(input_revision or {}), "input_sha256": hashlib.sha256(workbook_bytes).hexdigest(),
        "mass_balance": mass, "resources": resources,
        "inventory": {}, "policy_warnings": list(analysis.policy_warnings), "status": status,
        "publish_status": "ready_for_publish" if all(part["ok"] for part in status.values()) else "review_required",
    }


def prepare_weekly_schedule_update(workbook_bytes: bytes, *, plan_year: int, plan_month: int, input_revision=None):
    analysis = analyze_weekly_workbook(workbook_bytes, plan_year=plan_year, plan_month=plan_month)
    updated = patch_weekly_workbook(workbook_bytes, analysis)
    report = build_weekly_schedule_report(workbook_bytes, analysis, input_revision=input_revision)
    return updated, report, analysis


def verify_weekly_workbook(workbook_bytes: bytes, *, schedule_report: dict[str, Any], plan_year: int, plan_month: int):
    revision = schedule_report.get("input_revision")
    if not isinstance(revision, dict) or not revision:
        raise RuntimeError("Weekly schedule report thiếu input_revision/provenance.")
    if schedule_report.get("plan_month") != f"{plan_year:04d}-{plan_month:02d}":
        raise RuntimeError("Weekly schedule report sai kỳ kế hoạch.")
    analysis = analyze_weekly_workbook(workbook_bytes, plan_year=plan_year, plan_month=plan_month)
    if analysis.changed_cells:
        raise RuntimeError(f"Workbook chưa khớp weekly engine: {analysis.changed_cells} ô sai.")
    expected_mass, carryovers = _mass_balance(analysis)
    actual_mass = schedule_report.get("mass_balance") or {}
    for code, expected in expected_mass.items():
        actual = actual_mass.get(code) or {}
        for key in ("planned_qty", "scheduled_qty", "carryover_qty"):
            if not math.isclose(float(actual.get(key, 0) or 0), float(expected[key]), rel_tol=1e-9, abs_tol=BALANCE_EPS):
                raise RuntimeError(f"Mass balance {code}.{key} không khớp workbook.")
    expected_warnings = sorted((w["code"], w["type"]) for w in analysis.policy_warnings)
    report_warnings = sorted((str(w.get("code")), str(w.get("type"))) for w in schedule_report.get("policy_warnings", []) if isinstance(w, dict))
    if expected_warnings != report_warnings:
        raise RuntimeError("policy_warnings không khớp workbook.")
    expected_resource = _validate_shared_machine(analysis)
    report_resource = (schedule_report.get("status") or {}).get("resource_validation") or {}
    if (
        bool(report_resource.get("ok")) != bool(expected_resource["ok"])
        or report_resource.get("state") != expected_resource["state"]
    ):
        raise RuntimeError("resource_validation không khớp shared-machine schedule.")
    return {"validated": True, "algorithm": "ke_hoach_sx_tuan_v1", "checked_skus": len(analysis.calculated), "carryover_skus": carryovers, "policy_warning_count": len(analysis.policy_warnings), "publish_status": schedule_report.get("publish_status")}
