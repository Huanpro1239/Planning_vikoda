from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from planning_schedule_report import load_schedule_report
from sharepoint.client import GraphClient, get_access_token
from sync_stock import DEST_PATH

from .workbook import (
    PLANNING_SHEET,
    START_COLUMN,
    STOCK_SHEET,
    validate_workbook_context,
)

from .row_mass_balance import validate_planning_rows

from .shared_machine import (
    SHARED_LINES,
    SHARED_RESOURCE,
    _find_shared_resource_info,
    _header_day,
    _validate_report_provenance,
    _validate_shared_from_workbook,
    _validate_shared_timeline,
)

def verify_workbook(workbook_bytes, schedule_report=None, plan_year=None):
    workbook = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        workbook_context = validate_workbook_context(
            workbook,
            schedule_report=schedule_report,
            plan_year=plan_year,
        )
        planning = workbook_context["planning"]
        headers = workbook_context["headers"]
        plan_year = workbook_context["plan_year"]
        plan_month = workbook_context["plan_month"]
        stock_checked = workbook_context["stock_checked"]
        selector = workbook_context["selector"]
        source_column = workbook_context["source_column"]
        fc_checked = workbook_context["fc_checked"]
        expected_start = workbook_context["expected_start"]
        expected_end = workbook_context["expected_end"]

        # 4) Hậu kiểm độc lập dữ liệu kế hoạch và lịch ngày.
        EPS = 1e-7
        resource_capacity = {}
        daily_resource_usage = {}
        checked_schedule_rows = 0
        planning_data = {
            "__plan_year": plan_year,
            "__plan_month": plan_month,
        }

        validated_rows = validate_planning_rows(
            planning,
            headers,
            schedule_report=schedule_report,
        )
        for validated_row in validated_rows:
            code = validated_row["code"]
            line = validated_row["line"]
            per_shift = validated_row["per_shift"]
            shifts_per_day = validated_row["shifts_per_day"]
            daily_values = validated_row["daily_values"]
            planned = validated_row["planned"]
            total_scheduled = validated_row["scheduled"]

            resource = SHARED_RESOURCE if line in SHARED_LINES else line
            if resource:
                existing_capacity = resource_capacity.get(resource)
                if existing_capacity is None:
                    resource_capacity[resource] = shifts_per_day
                elif resource == SHARED_RESOURCE or line not in {"RGB", "Galon"}:
                    resource_capacity[resource] = min(existing_capacity, shifts_per_day)
                else:
                    resource_capacity[resource] = max(existing_capacity, shifts_per_day)

                if per_shift > EPS:
                    for current_day, qty in zip(headers, daily_values):
                        key = (resource, current_day)
                        daily_resource_usage[key] = daily_resource_usage.get(key, 0.0) + qty / per_shift

            planning_data[code] = {
                "line": line,
                "per_shift": per_shift,
                "daily_values": daily_values,
                "planned": planned,
                "scheduled": total_scheduled,
            }
            checked_schedule_rows += 1

        for (resource, current_day), used_shifts in daily_resource_usage.items():
            capacity = resource_capacity.get(resource, 0.0)
            if used_shifts > capacity + 1e-6:
                raise RuntimeError(
                    f"Resource {resource} ngày {_header_day(current_day)} có ít nhất "
                    f"{used_shifts:.3f} ca sản xuất > capacity {capacity:.3f}; chưa tính setup."
                )

        shared_codes = [
            code
            for code, item in planning_data.items()
            if not code.startswith("__")
            and item["line"] in SHARED_LINES
            and item["scheduled"] > EPS
        ]
        if len(shared_codes) > 1:
            shared_capacity = resource_capacity[SHARED_RESOURCE]
            _, shared_meta = _find_shared_resource_info(schedule_report)
            has_timeline = isinstance((shared_meta or {}).get("timeline"), list) and (
                shared_meta or {}
            ).get("timeline")
            if has_timeline:
                # Report có timeline production/setup: kiểm chứng nghiêm ngặt
                # không chồng lấn, đủ 0,5 ca setup khi đổi mã, và khớp workbook.
                _validate_shared_timeline(
                    schedule_report,
                    headers,
                    planning_data,
                    shared_capacity,
                )
            else:
                # Report hiện hành chưa xuất timeline: hậu kiểm tự lập từ workbook
                # (ràng buộc capacity/ngày đã kiểm ở trên; ở đây kiểm chặn dưới
                # theo tháng có tính setup).
                _validate_shared_from_workbook(
                    headers,
                    planning_data,
                    shared_capacity,
                )

        if schedule_report is not None:
            _validate_report_provenance(schedule_report, plan_year, plan_month)

        print(
            f"[VERIFY] {stock_checked} mã J=Ton_kho!D, K=SUM(Ton_kho!E:H) đúng; "
            f"{selector!r} -> FC!{get_column_letter(source_column)}, {fc_checked} mã L đúng; "
            f"lịch {expected_start.splitlines()[0]} -> {expected_end.splitlines()[0]} đúng."
        )
        return {
            "selector": selector,
            "plan_month": plan_month,
            "plan_year": plan_year,
            "source_column": source_column,
            "stock_checked": stock_checked,
            "fc_checked": fc_checked,
            "schedule_checked": checked_schedule_rows,
            "publish_status": (
                schedule_report.get("publish_status")
                if isinstance(schedule_report, dict)
                else "verified_without_carryover_report"
            ),
        }
    finally:
        workbook.close()


def main():
    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)
    dest_item = graph.get_item_by_path(drive_id, DEST_PATH)
    dest_bytes = graph.download_file(drive_id, dest_item["id"])
    verify_workbook(dest_bytes, schedule_report=load_schedule_report())


if __name__ == "__main__":
    main()
