import hashlib
import json
from datetime import date, datetime
from pathlib import Path


REPORT_PATH = Path("planning_schedule_report.json")


def _json_value(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(key): _json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_value(item) for item in value]
    return value


def json_safe(value):
    """Return a recursively JSON-serializable copy of planning metadata."""
    return _json_value(value)


def attach_output_hash(report, workbook_bytes):
    result = dict(report)
    result["output_sha256"] = hashlib.sha256(workbook_bytes).hexdigest()
    return result


def save_schedule_report(report, path=REPORT_PATH):
    path = Path(path)
    path.write_text(
        json.dumps(json_safe(report), ensure_ascii=False, indent=2, sort_keys=True)
        + "\n",
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
        f"plan={report.get('plan_month')}; "
        f"status={report.get('publish_status')}."
    )
    status = report.get("status") or {}
    if status:
        print(
            "[PLAN_REPORT][STATUS] "
            f"monthly={((status.get('monthly_quantity') or {}).get('state'))}; "
            f"resource={((status.get('resource_validation') or {}).get('state'))}; "
            f"service={((status.get('service') or {}).get('state'))}; "
            f"safety_stock={((status.get('safety_stock') or {}).get('state'))}."
        )

    service_short = []
    buffer_short = []
    for code, values in sorted((report.get("mass_balance") or {}).items()):
        service_missing = float(values.get("service_carryover_qty", 0) or 0)
        buffer_missing = float(values.get("buffer_carryover_qty", 0) or 0)
        if service_missing > 1e-9:
            service_short.append((code, service_missing, values.get("uom") or ""))
        if buffer_missing > 1e-9:
            buffer_short.append((code, buffer_missing, values.get("uom") or ""))

    if service_short:
        print("[PLAN_REPORT][SERVICE_SHORTFALL] Chưa đủ sản lượng phục vụ bán hàng:")
        for code, qty, uom in service_short:
            print(f"  - {code}: {qty:g} {uom}")
    if buffer_short:
        print("[PLAN_REPORT][SAFETY_STOCK_SHORTFALL] Chưa đạt tồn cuối dự kiến:")
        for code, qty, uom in buffer_short:
            print(f"  - {code}: {qty:g} {uom}")
