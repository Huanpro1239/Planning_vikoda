import json
import math
import time
from pathlib import Path

from graph_retry import install_retry_after_support, retry_delay_seconds as retry_wait_seconds
from planning_pipeline import prepare_pipeline_output
from planning_schedule_report import print_operational_report
import sync_planning_metrics as metrics
import sync_planning_schedule as schedule_base
import sync_stock
from sync_stock import GraphClient, get_access_token, is_retryable_graph_error


OUTPUT_DIR = Path("dry_run_artifacts")
FOCUS_CODES = ("130100006", "130100008", "130100013")

SOURCES = {
    "actual_stock": sync_stock.SOURCE_ACTUAL_PATH,
    "factory_vikoda": sync_stock.SOURCE_FACTORY_VIKODA_PATH,
    "factory_vkd": sync_stock.SOURCE_FACTORY_VKD_PATH,
    "accounting_vikoda": sync_stock.SOURCE_ACCOUNTING_VIKODA_PATH,
    "accounting_vkd": sync_stock.SOURCE_ACCOUNTING_VKD_PATH,
}


def _read_item_with_retry(
    graph,
    drive_id,
    path,
    *,
    sleep_func=time.sleep,
    max_attempts=6,
    retry_delay_seconds=5,
):
    for attempt in range(1, max_attempts + 1):
        try:
            item = graph.get_item_by_path(drive_id, path)
            data = graph.download_file(drive_id, item["id"])
            return item, data
        except Exception as exc:
            if not is_retryable_graph_error(exc) or attempt == max_attempts:
                raise
            delay = retry_wait_seconds(exc, retry_delay_seconds)
            print(
                f"[DRY-RUN][READ] {path}: lỗi tạm thời; thử lại "
                f"({attempt}/{max_attempts}) sau {delay:g}s."
            )
            sleep_func(delay)
    raise RuntimeError(f"Không đọc được {path}.")


def _existing_schedule_analysis(workbook_bytes, plan_year):
    headers, products = schedule_base.read_schedule_inputs(
        workbook_bytes,
        plan_year=plan_year,
    )
    schedule = schedule_base._empty_schedule(headers, products)
    mass_balance = {}
    shared_usage = {day: 0.0 for day in headers}

    for product in products:
        code = product["code"]
        for index, day in enumerate(headers):
            raw = product["existing_daily"][index]
            qty = 0.0 if raw in (None, "") else float(raw)
            schedule[code][day] = qty
            if product["line"] in {"KHS", "PET 9000"} and product["per_shift"] > 0:
                shared_usage[day] += qty / product["per_shift"]

        planned = float(product["planned_qty"] or 0)
        scheduled = sum(schedule[code].values())
        mass_balance[code] = {
            "planned_qty": planned,
            "scheduled_qty": scheduled,
            "delta_p_minus_scheduled": planned - scheduled,
        }

    inventory = schedule_base.simulate_inventory(headers, products, schedule)
    return {
        "mass_balance_mismatch": {
            code: row
            for code, row in mass_balance.items()
            if not math.isclose(
                row["planned_qty"],
                row["scheduled_qty"],
                rel_tol=1e-9,
                abs_tol=1e-5,
            )
        },
        "stockout": {
            code: values["first_stockout"].isoformat()
            for code, values in inventory.items()
            if values.get("first_stockout") is not None
        },
        "below_safety": {
            code: values["first_below_safety"].isoformat()
            for code, values in inventory.items()
            if values.get("first_below_safety") is not None
        },
        "shared_machine_production_only_max_shifts": max(shared_usage.values()) if shared_usage else 0,
        "shared_machine_setup_verifiable": False,
        "note": (
            "Workbook hiện hành chỉ có qty/ngày; không có interval/setup provenance, "
            "nên không thể chứng minh capacity máy chung có tính setup."
        ),
    }


def _focus_daily_analysis(workbook_bytes, plan_year):
    headers, products = schedule_base.read_schedule_inputs(
        workbook_bytes,
        plan_year=plan_year,
    )
    by_code = {product["code"]: product for product in products}
    result = {}

    for code in FOCUS_CODES:
        product = by_code.get(code)
        if product is None:
            result[code] = {"missing": True, "days": []}
            continue

        balance = float(product.get("actual_stock", 0) or 0)
        days = []
        for index, current_day in enumerate(headers):
            raw = product["existing_daily"][index]
            production = 0.0 if raw in (None, "") else float(raw)
            demand = float(product["demand_by_day"].get(current_day, 0) or 0)
            debt_due = float(product.get("debt", 0) or 0) if index == 0 else 0.0
            opening_net = balance
            balance += production - demand - debt_due
            days.append(
                {
                    "date": current_day.isoformat(),
                    "opening_net": opening_net,
                    "production": production,
                    "demand": demand,
                    "debt_due": debt_due,
                    "ending_net": balance,
                    "stockout": balance < -1e-7,
                }
            )

        result[code] = {
            "line": product.get("line"),
            "uom": product.get("uom"),
            "actual_stock": product.get("actual_stock"),
            "debt": product.get("debt"),
            "planned_qty": product.get("planned_qty"),
            "days": days,
        }
    return result


def _classify_late_completed(report):
    buffer_only = []
    demand_late = []
    indeterminate = []

    for resource, values in (report.get("resources") or {}).items():
        meta = values.get("meta") or {}
        for raw in meta.get("late_completed") or []:
            event = dict(raw)
            event["resource"] = resource
            finish = event.get("finish_shift")
            demand_deadline = event.get("demand_deadline_shift")
            try:
                finish_value = float(finish)
                demand_value = float(demand_deadline)
            except (TypeError, ValueError):
                indeterminate.append(event)
                continue

            if not math.isfinite(finish_value) or not math.isfinite(demand_value):
                indeterminate.append(event)
            elif finish_value <= demand_value + 1e-7:
                # Missed the one-day production buffer, but still completed by
                # the business demand due point.
                buffer_only.append(event)
            else:
                demand_late.append(event)

    return {
        "buffer_only": buffer_only,
        "demand_late": demand_late,
        "indeterminate": indeterminate,
    }


def _after_summary(report, final_bytes, plan_year):
    carryover = {
        code: values
        for code, values in (report.get("mass_balance") or {}).items()
        if float(values.get("carryover_qty", 0) or 0) > 1e-9
    }
    mass_errors = {
        code: values
        for code, values in (report.get("mass_balance") or {}).items()
        if not math.isclose(
            float(values.get("balanced_qty", 0) or 0),
            float(values.get("planned_qty", 0) or 0),
            rel_tol=1e-9,
            abs_tol=1e-5,
        )
    }
    stockout = {
        code: values.get("first_stockout")
        for code, values in (report.get("inventory") or {}).items()
        if values.get("first_stockout")
    }
    below_safety = {
        code: values.get("first_below_safety")
        for code, values in (report.get("inventory") or {}).items()
        if values.get("first_below_safety")
    }

    resources = {}
    for name, values in (report.get("resources") or {}).items():
        meta = values.get("meta") or {}
        timeline = meta.get("timeline") or []
        setup_shifts = float(meta.get("setup_shifts", 0) or 0)
        resources[name] = {
            "capacity_shifts_per_day": values.get("capacity_shifts_per_day"),
            "utilization": values.get("utilization"),
            "setup_shifts": setup_shifts,
            "timeline_events": len(timeline),
            "unserved_due": len(meta.get("unserved_due") or []),
            "late_completed": len(meta.get("late_completed") or []),
            "setup_independently_verifiable": bool(timeline) or setup_shifts <= 1e-9,
            "setup_verification_note": (
                "timeline provenance available"
                if timeline
                else (
                    "no setup required"
                    if setup_shifts <= 1e-9
                    else "setup_shifts reported but no independent timeline provenance"
                )
            ),
        }

    return {
        "publish_status": report.get("publish_status"),
        "mass_balance_errors": mass_errors,
        "carryover": carryover,
        "stockout": stockout,
        "below_safety": below_safety,
        "resources": resources,
        "focus_daily": _focus_daily_analysis(final_bytes, plan_year),
        "late_classification": _classify_late_completed(report),
    }


def _summary_markdown(summary):
    before = summary["before"]
    after = summary["after"]
    late = after["late_classification"]
    lines = [
        "# Vikoda planning dry-run",
        "",
        f"- Target ETag snapshot: `{summary['input_revision']['target']['etag']}`",
        f"- Plan: `{summary['report']['plan_month']}`",
        f"- Algorithm: `{summary['report']['algorithm']}`",
        "- SharePoint writes: **0**",
        "",
        "## Before (workbook copy as read)",
        f"- Mass-balance mismatches: {len(before['mass_balance_mismatch'])}",
        f"- Stockout-risk SKUs: {len(before['stockout'])}",
        f"- Below-safety SKUs: {len(before['below_safety'])}",
        f"- Shared machine max production-only shifts/day: {before['shared_machine_production_only_max_shifts']:.3f}",
        "- Setup/timeline proof: unavailable in old workbook cells",
        "",
        "## After (full in-memory pipeline)",
        f"- Publish status: `{after['publish_status']}`",
        f"- Mass-balance errors: {len(after['mass_balance_errors'])}",
        f"- Carryover SKUs: {len(after['carryover'])}",
        f"- Stockout-risk SKUs: {len(after['stockout'])}",
        f"- Below-safety SKUs: {len(after['below_safety'])}",
        f"- Late vs production buffer only: {len(late['buffer_only'])}",
        f"- Late vs actual demand due: {len(late['demand_late'])}",
        f"- Late classification indeterminate: {len(late['indeterminate'])}",
        "",
        "## Resources",
    ]
    for name, values in after["resources"].items():
        util = 100 * float(values.get("utilization", 0) or 0)
        lines.append(
            f"- **{name}**: capacity={values.get('capacity_shifts_per_day')} ca/ngày; "
            f"utilization={util:.1f}%; setup={values.get('setup_shifts')} ca; "
            f"timeline={values.get('timeline_events')} events; "
            f"unserved_due={values.get('unserved_due')}; late={values.get('late_completed')}; "
            f"setup_provenance={'YES' if values.get('setup_independently_verifiable') else 'NO'} "
            f"({values.get('setup_verification_note')})"
        )

    if after["carryover"]:
        lines.extend(["", "## Carryover"])
        for code, values in after["carryover"].items():
            lines.append(
                f"- {code}: {values.get('carryover_qty')} {values.get('uom') or ''} "
                f"(scheduled={values.get('scheduled_qty')}, P={values.get('planned_qty')})"
            )
    if after["stockout"]:
        lines.extend(["", "## Stockout risk"])
        for code, day in after["stockout"].items():
            lines.append(f"- {code}: {day}")

    lines.extend(["", "## Focus SKU daily balance"])
    for code in FOCUS_CODES:
        detail = after["focus_daily"].get(code, {})
        if detail.get("missing"):
            lines.append(f"- {code}: missing from planning workbook")
            continue
        lines.append(
            f"### {code} — line={detail.get('line')}, J={detail.get('actual_stock')}, "
            f"N={detail.get('debt')}, P={detail.get('planned_qty')} {detail.get('uom') or ''}"
        )
        lines.append("| Date | Opening net | Production | Demand | Debt due | Ending net | Stockout |")
        lines.append("|---|---:|---:|---:|---:|---:|:---:|")
        for day in detail.get("days", []):
            lines.append(
                f"| {day['date']} | {day['opening_net']:.3f} | {day['production']:.3f} | "
                f"{day['demand']:.3f} | {day['debt_due']:.3f} | {day['ending_net']:.3f} | "
                f"{'YES' if day['stockout'] else ''} |"
            )

    lines.extend(["", "## Late quantum classification"])
    for label, title in (
        ("buffer_only", "Missed production buffer but met actual demand due"),
        ("demand_late", "Completed after actual demand due"),
        ("indeterminate", "Indeterminate"),
    ):
        lines.append(f"### {title}: {len(late[label])}")
        for event in late[label]:
            lines.append(
                f"- {event.get('resource')} {event.get('code')} unit {event.get('unit_no')}: "
                f"qty={event.get('qty')} {event.get('uom') or ''}; "
                f"production_deadline={event.get('production_deadline_date')}; "
                f"demand_due={event.get('demand_due_date')}; finish_shift={event.get('finish_shift')}"
            )

    changes = summary["report"].get("pipeline", {}).get("planning_changes", {})
    lines.extend(["", "## J:R changes on the dry-run copy", f"- Changed SKUs: {len(changes)}"])
    for code, fields in list(changes.items())[:20]:
        details = ", ".join(
            f"{field}: {values.get('before')} → {values.get('after')}"
            for field, values in fields.items()
        )
        lines.append(f"- {code}: {details}")

    return "\n".join(lines) + "\n"


def main():
    install_retry_after_support()
    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    # READ ONLY from SharePoint. This script has no upload_file call.
    target_item, target_bytes = _read_item_with_retry(
        graph,
        drive_id,
        sync_stock.DEST_PATH,
    )
    source_bytes = {}
    source_revision = {}
    for key, path in SOURCES.items():
        item, data = _read_item_with_retry(graph, drive_id, path)
        source_bytes[key] = data
        source_revision[key] = {
            "path": path,
            "etag": item.get("eTag"),
            "last_modified": item.get("lastModifiedDateTime"),
        }

    report_date, _, _ = metrics.read_actual_inputs(source_bytes["actual_stock"])
    # Resolve plan year for the BEFORE read consistently with raw-source date.
    _, plan_month, _ = metrics.read_planning_rows(target_bytes)
    plan_year = metrics._resolve_plan_year(report_date, plan_month)
    before = _existing_schedule_analysis(target_bytes, plan_year)

    revision = {
        "target": {
            "path": sync_stock.DEST_PATH,
            "item_id": target_item.get("id"),
            "etag": target_item.get("eTag"),
            "last_modified": target_item.get("lastModifiedDateTime"),
        },
        "sources": source_revision,
    }

    final_bytes, report, proposed_state = prepare_pipeline_output(
        target_bytes,
        source_bytes,
        runtime_state=metrics.load_runtime_state(),
        input_revision=revision,
    )
    after = _after_summary(report, final_bytes, plan_year)

    summary = {
        "mode": "read_only_sharepoint_full_pipeline_dry_run",
        "sharepoint_writes": 0,
        "input_revision": revision,
        "before": before,
        "after": after,
        "report": report,
    }

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUTPUT_DIR / "Sap_ke_hoach_dry_run.xlsx").write_bytes(final_bytes)
    (OUTPUT_DIR / "planning_schedule_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    (OUTPUT_DIR / "dry_run_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    (OUTPUT_DIR / "proposed_runtime_state.json").write_text(
        json.dumps(proposed_state, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    (OUTPUT_DIR / "dry_run_summary.md").write_text(
        _summary_markdown(summary),
        encoding="utf-8",
    )

    print_operational_report(report)
    print(_summary_markdown(summary))
    print("[DRY-RUN] Hoàn tất. Không có lệnh upload SharePoint trong script.")


if __name__ == "__main__":
    main()
