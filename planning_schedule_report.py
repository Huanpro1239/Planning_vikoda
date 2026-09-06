import hashlib
import json
import math
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook


REPORT_PATH = Path("planning_schedule_report.json")
REPORT_SCHEMA_VERSION = 2
PLANNING_SHEET = "Ke_hoach_SX"
MASS_BALANCE_EPS = 1e-5


def _json_value(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(key): _json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_value(item) for item in value]
    return value


def _read_uom_by_code(workbook_bytes):
    workbook = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        if PLANNING_SHEET not in workbook.sheetnames:
            return {}
        sheet = workbook[PLANNING_SHEET]
        result = {}
        for row in range(2, sheet.max_row + 1):
            raw_code = sheet.cell(row=row, column=1).value
            if raw_code in (None, ""):
                continue
            if isinstance(raw_code, float) and raw_code.is_integer():
                raw_code = int(raw_code)
            code = str(raw_code).strip()
            result[code] = str(sheet.cell(row=row, column=3).value or "").strip()
        return result
    finally:
        workbook.close()


def _validate_mass_balance_values(code, planned, scheduled, carryover):
    values = {
        "planned_qty": float(planned),
        "scheduled_qty": float(scheduled),
        "carryover_qty": float(carryover),
    }
    for label, value in values.items():
        if not math.isfinite(value):
            raise RuntimeError(
                f"Schedule report mã {code} có {label} không hữu hạn: {value!r}."
            )
        if value < -MASS_BALANCE_EPS:
            raise RuntimeError(
                f"Schedule report mã {code} có {label} âm: {value}."
            )

    planned = values["planned_qty"]
    scheduled = values["scheduled_qty"]
    carryover = values["carryover_qty"]
    if scheduled > planned + MASS_BALANCE_EPS:
        raise RuntimeError(
            f"Schedule report mã {code} scheduled={scheduled} vượt P={planned}."
        )
    if carryover > planned + MASS_BALANCE_EPS:
        raise RuntimeError(
            f"Schedule report mã {code} carryover={carryover} vượt P={planned}."
        )
    if not math.isclose(
        scheduled + carryover,
        planned,
        rel_tol=1e-9,
        abs_tol=MASS_BALANCE_EPS,
    ):
        raise RuntimeError(
            f"Schedule report mã {code} mất cân đối: scheduled {scheduled} + "
            f"carryover {carryover} != P {planned}."
        )
    return planned, scheduled, carryover


def build_schedule_report(
    workbook_bytes,
    info,
    *,
    input_revision=None,
    algorithm="priority_v8",
):
    """Build an operational report that can independently prove mass balance."""
    headers = list(info["headers"])
    products = list(info["products"])
    schedule = info["schedule"]
    carryover = dict(info.get("carryover") or {})
    utilization = dict(info.get("utilization") or {})
    uom_by_code = _read_uom_by_code(workbook_bytes)

    mass_balance = {}
    for product in products:
        code = product["code"]
        planned = float(product.get("planned_qty", 0) or 0)
        scheduled = sum(float(schedule[code].get(day, 0) or 0) for day in headers)
        missing = float(carryover.get(code, 0) or 0)
        planned, scheduled, missing = _validate_mass_balance_values(
            code,
            planned,
            scheduled,
            missing,
        )
        mass_balance[code] = {
            "uom": str(product.get("uom") or uom_by_code.get(code, "")),
            "planned_qty": planned,
            "scheduled_qty": scheduled,
            "carryover_qty": missing,
            "balanced_qty": scheduled + missing,
        }

    resources = {}
    for resource, capacity in sorted((info.get("line_capacity") or {}).items()):
        meta = dict((info.get("optimizer_meta") or {}).get(resource, {}) or {})
        for field in ("unserved_due", "late_completed", "deadline_misses", "shortage_by_day", "timeline"):
            records = meta.get(field)
            if isinstance(records, list):
                enriched = []
                for record in records:
                    if not isinstance(record, dict):
                        enriched.append(record)
                        continue
                    item = dict(record)
                    code = str(item.get("code") or "")
                    if code and not item.get("uom"):
                        item["uom"] = uom_by_code.get(code, "")
                    enriched.append(item)
                meta[field] = enriched
        resources[resource] = {
            "capacity_shifts_per_day": float(capacity),
            "utilization": float(utilization.get(resource, 0) or 0),
            "meta": meta,
        }

    inventory = {}
    for code, values in (info.get("inventory") or {}).items():
        inventory[code] = dict(values)

    report = {
        "schema_version": REPORT_SCHEMA_VERSION,
        "algorithm": algorithm,
        "plan_month": headers[0].strftime("%Y-%m") if headers else None,
        "input_revision": dict(input_revision or {}),
        "input_sha256": hashlib.sha256(workbook_bytes).hexdigest(),
        "mass_balance": mass_balance,
        "resources": resources,
        "inventory": inventory,
        "publish_status": (
            "feasible" if not carryover else "valid_but_infeasible"
        ),
    }
    return _json_value(report)


def attach_output_hash(report, workbook_bytes):
    result = dict(report)
    result["output_sha256"] = hashlib.sha256(workbook_bytes).hexdigest()
    return result


def save_schedule_report(report, path=REPORT_PATH):
    path = Path(path)
    path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    return path


def load_schedule_report(path=REPORT_PATH):
    path = Path(path)
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def print_operational_report(report):
    print(
        f"[PLAN_REPORT] algorithm={report.get('algorithm')}; "
        f"plan={report.get('plan_month')}; status={report.get('publish_status')}."
    )

    carryovers = [
        (code, values)
        for code, values in sorted((report.get("mass_balance") or {}).items())
        if float(values.get("carryover_qty", 0) or 0) > 1e-9
    ]
    if carryovers:
        print("[PLAN_REPORT][CARRYOVER] Sản lượng hợp lệ nhưng chưa xếp được:")
        for code, values in carryovers:
            print(
                f"  - {code}: {values['carryover_qty']:g} {values.get('uom') or ''}; "
                f"scheduled={values['scheduled_qty']:g}; P={values['planned_qty']:g}"
            )

    for resource, resource_info in sorted((report.get("resources") or {}).items()):
        meta = resource_info.get("meta") or {}
        for event in meta.get("unserved_due") or []:
            print(
                f"[PLAN_REPORT][UNSERVED_DUE] {resource} {event.get('code')}: "
                f"{event.get('qty')} {event.get('uom') or ''}; "
                f"need={event.get('demand_due_date')}; "
                f"production_deadline={event.get('production_deadline_date')}"
            )
        for event in meta.get("late_completed") or []:
            print(
                f"[PLAN_REPORT][LATE_COMPLETED] {resource} {event.get('code')}: "
                f"{event.get('qty')} {event.get('uom') or ''}; "
                f"need={event.get('demand_due_date')}; "
                f"production_deadline={event.get('production_deadline_date')}"
            )
