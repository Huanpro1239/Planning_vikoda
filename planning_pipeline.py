import copy
import math
from io import BytesIO

from openpyxl import load_workbook

from planning_cleanup import remove_sheet_formulas
from planning_schedule_report import (
    attach_output_hash,
    build_schedule_report,
    json_safe,
)
from sync_planning_calendar_all_months import prepare_calendar_update_all_months
from sync_planning_fc import prepare_planning_fc_update
import sync_planning_metrics as metrics
import sync_planning_metrics_all_months  # noqa: F401 - installs all-month/direct hooks
import sync_planning_metrics_compat as metrics_compat
from sync_planning_schedule_production import install_production_output_cleanup
import sync_planning_schedule_priority as priority
from sync_planning_stock_inputs import prepare_stock_input_update
import sync_stock
from sync_stock_compat import read_conversion_factors_robust
from verify_planning_month import verify_workbook


SOURCE_KEYS = (
    "actual_stock",
    "factory_vikoda",
    "factory_vkd",
    "accounting_vikoda",
    "accounting_vkd",
)


def _planning_snapshot(workbook_bytes):
    workbook = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        sheet = workbook["Ke_hoach_SX"]
        result = {}
        for row in range(2, sheet.max_row + 1):
            raw_code = sheet.cell(row=row, column=1).value
            code = sync_stock.normalize_code(raw_code)
            if not code:
                continue
            result[code] = {
                "J": sheet.cell(row=row, column=10).value,
                "K": sheet.cell(row=row, column=11).value,
                "L": sheet.cell(row=row, column=12).value,
                "M": sheet.cell(row=row, column=13).value,
                "N": sheet.cell(row=row, column=14).value,
                "O": sheet.cell(row=row, column=15).value,
                "P": sheet.cell(row=row, column=16).value,
                "Q": sheet.cell(row=row, column=17).value,
                "R": sheet.cell(row=row, column=18).value,
            }
        return result
    finally:
        workbook.close()


def _diff_snapshot(before, after):
    changed = {}
    for code in sorted(set(before) | set(after)):
        fields = {}
        for field in ("J", "K", "L", "M", "N", "O", "P", "Q", "R"):
            old = before.get(code, {}).get(field)
            new = after.get(code, {}).get(field)
            same = old == new
            if isinstance(old, (int, float)) and isinstance(new, (int, float)):
                same = math.isclose(float(old), float(new), rel_tol=1e-10, abs_tol=1e-7)
            if not same:
                fields[field] = {"before": old, "after": new}
        if fields:
            changed[code] = fields
    return changed


def prepare_pipeline_output(
    target_workbook_bytes,
    source_bytes,
    *,
    runtime_state,
    input_revision=None,
):
    """
    Compute the complete planning output on one in-memory workbook snapshot.

    This function performs no network calls and no file writes. The caller may
    persist the returned bytes only after validation succeeds.
    """
    missing_sources = [key for key in SOURCE_KEYS if key not in source_bytes]
    if missing_sources:
        raise RuntimeError("Thiếu nguồn dry-run: " + ", ".join(missing_sources))

    install_production_output_cleanup()

    initial_bytes = target_workbook_bytes
    before_snapshot = _planning_snapshot(initial_bytes)
    work = initial_bytes
    steps = []

    # 1) Recompute Ton_kho from the same raw source snapshots production uses.
    conversion_factors, conversion_hash = read_conversion_factors_robust(work)
    actual_stock = sync_stock.read_actual_stock(source_bytes["actual_stock"])
    factory_vikoda = sync_stock.read_single_value_source(
        source_bytes["factory_vikoda"],
        label="Tồn nhà máy Vikoda",
        source_name="NXT_Vikoda.xlsm",
        sheet_name="Sheet1",
        code_column=2,
        value_column=12,
        value_column_letter="L",
    )
    factory_vkd = sync_stock.read_single_value_source(
        source_bytes["factory_vkd"],
        label="Tồn nhà máy VKD",
        source_name="NXT_VKD.xlsm",
        sheet_name="Sheet1",
        code_column=2,
        value_column=12,
        value_column_letter="L",
        vkd_to_vikoda=True,
    )
    accounting_vikoda = sync_stock.read_single_value_source(
        source_bytes["accounting_vikoda"],
        label="Tồn kế toán Vikoda",
        source_name="XNT_ketoan_Vikoda.xlsm",
        sheet_name="Sheet1",
        code_column=2,
        value_column=13,
        value_column_letter="M",
    )
    accounting_vkd = sync_stock.read_single_value_source(
        source_bytes["accounting_vkd"],
        label="Tồn kế toán VKD",
        source_name="XNT_ketoan_VKD.xlsm",
        sheet_name="Sheet1",
        code_column=2,
        value_column=13,
        value_column_letter="M",
        vkd_to_vikoda=True,
    )
    work = sync_stock.patch_destination_workbook(
        work,
        actual_stock=actual_stock,
        factory_vikoda=factory_vikoda,
        factory_vkd=factory_vkd,
        accounting_vikoda=accounting_vikoda,
        accounting_vkd=accounting_vkd,
        conversion_factors=conversion_factors,
    )
    steps.append("Ton_kho")

    # 2) Refresh J/K from Ton_kho and L from selected FC month.
    work, stock_info = prepare_stock_input_update(work)
    work, fc_info = prepare_planning_fc_update(work)
    steps.extend(["J_K", "FC_L"])

    # 3) Resolve planning month/year from actual-stock report snapshot.
    report_date, actual_receipts, consignments = metrics.read_actual_inputs(
        source_bytes["actual_stock"]
    )
    selector, plan_month, planning_rows = metrics.read_planning_rows(work)
    plan_year = metrics._resolve_plan_year(report_date, plan_month)

    # 4) Calendar is computed using the same resolved plan year, not wall clock.
    work, calendar_info = prepare_calendar_update_all_months(
        work,
        plan_year=plan_year,
    )
    steps.append("calendar")

    # Re-read rows after calendar/layout migration and load master Leadtime.
    selector, plan_month, planning_rows = metrics.read_planning_rows(work)
    conversion_factors, _ = metrics_compat.read_conversion_factors_and_leadtime(work)
    system_receipts = metrics.read_system_receipts(
        source_bytes["factory_vikoda"],
        conversion_factors,
        set(planning_rows),
    )

    # 5) Calculate M:R from a copy of runtime state. Dry-run never mutates state.
    next_state = copy.deepcopy(runtime_state or {})
    metric_values, state_changed, source_key, plan_key = metrics.calculate_metrics(
        report_date=report_date,
        plan_month=plan_month,
        planning_rows=planning_rows,
        actual_receipts=actual_receipts,
        system_receipts=system_receipts,
        current_consignments=consignments,
        state=next_state,
    )
    metric_changes = metrics.count_changes(planning_rows, metric_values)
    if metric_changes:
        work, metric_patched = metrics.patch_workbook(work, metric_values)
    else:
        metric_patched = 0
    steps.append("M_R")

    # 6) Build V8 schedule on the same bytes; no upload occurs here.
    scheduled_bytes, schedule_info = priority.base.prepare_schedule_update(
        work,
        plan_year=plan_year,
    )
    steps.append("schedule")

    report = build_schedule_report(
        work,
        schedule_info,
        input_revision=dict(input_revision or {}),
        algorithm="priority_v8",
    )

    # 7) Cleanup formulas on the local copy, then validate exact final bytes.
    final_bytes, formulas_removed = remove_sheet_formulas(
        scheduled_bytes,
        "Ke_hoach_SX",
    )
    steps.append("cleanup")
    report = attach_output_hash(report, final_bytes)
    verify_info = verify_workbook(final_bytes, schedule_report=report)
    steps.append("verify")

    after_snapshot = _planning_snapshot(final_bytes)
    report["pipeline"] = {
        "mode": "offline_single_snapshot",
        "steps": steps,
        "source_period": source_key,
        "plan_period": plan_key,
        "selector": selector,
        "conversion_hash": conversion_hash,
        "state_would_change": bool(state_changed),
        "stock_rows_changed": stock_info.get("changed_count", 0),
        "fc_rows_changed": fc_info.get("changed_count", 0),
        "calendar_cells_changed": calendar_info.get("changed_count", 0),
        "metric_rows_changed": metric_changes,
        "metric_rows_patched": metric_patched,
        "schedule_cells_changed": schedule_info.get("changed_count", 0),
        "formulas_removed": formulas_removed,
        "planning_changes": _diff_snapshot(before_snapshot, after_snapshot),
        "verify": verify_info,
    }

    # The pipeline block is appended after build_schedule_report(), so normalize
    # it again here. This makes dry-run artifacts and post-upload audit saving
    # safe even when before/after R values are datetime objects.
    return final_bytes, json_safe(report), next_state
