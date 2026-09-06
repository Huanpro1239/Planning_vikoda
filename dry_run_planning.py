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


def _after_summary(report):
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
        resources[name] = {
            "capacity_shifts_per_day": values.get("capacity_shifts_per_day"),
            "utilization": values.get("utilization"),
            "setup_shifts": meta.get("setup_shifts", 0),
            "timeline_events": len(meta.get("timeline") or []),
            "unserved_due": len(meta.get("unserved_due") or []),
            "late_completed": len(meta.get("late_completed") or []),
        }

    return {
        "publish_status": report.get("publish_status"),
        "mass_balance_errors": mass_errors,
        "carryover": carryover,
        "stockout": stockout,
        "below_safety": below_safety,
        "resources": resources,
    }


def _summary_markdown(summary):
    before = summary["before"]
    after = summary["after"]
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
        "",
        "## Resources",
    ]
    for name, values in after["resources"].items():
        util = 100 * float(values.get("utilization", 0) or 0)
        lines.append(
            f"- **{name}**: capacity={values.get('capacity_shifts_per_day')} ca/ngày; "
            f"utilization={util:.1f}%; setup={values.get('setup_shifts')} ca; "
            f"timeline={values.get('timeline_events')} events; "
            f"unserved_due={values.get('unserved_due')}; late={values.get('late_completed')}"
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
    after = _after_summary(report)

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
