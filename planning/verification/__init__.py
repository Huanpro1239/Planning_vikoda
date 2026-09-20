import math
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from planning_schedule_report import load_schedule_report
from sharepoint.client import GraphClient, get_access_token
from sync_stock import DEST_PATH, normalize_code
from planning.weekly_engine import is_sugar_classification

from .workbook import (
    PLANNING_SHEET,
    START_COLUMN,
    STOCK_SHEET,
    validate_workbook_context,
)

from .shared_machine import (
    SHARED_LINES,
    SHARED_RESOURCE,
    _find_shared_resource_info,
    _finite_number,
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
        MASS_EPS = 1e-5
        resource_capacity = {}
        daily_resource_usage = {}
        checked_schedule_rows = 0
        planning_data = {
            "__plan_year": plan_year,
            "__plan_month": plan_month,
        }

        for row in range(2, planning.max_row + 1):
            code = normalize_code(planning.cell(row=row, column=1).value)
            if not code:
                continue

            batch = _finite_number(planning.cell(row=row, column=4).value, f"D{row}")
            per_shift = _finite_number(planning.cell(row=row, column=5).value, f"E{row}")
            line = str(planning.cell(row=row, column=6).value or "").strip()
            classification = str(planning.cell(row=row, column=8).value or "").strip()
            shifts_per_day = _finite_number(planning.cell(row=row, column=9).value, f"I{row}")
            _finite_number(planning.cell(row=row, column=10).value, f"J{row}")
            book_stock = _finite_number(planning.cell(row=row, column=11).value, f"K{row}")
            forecast = _finite_number(planning.cell(row=row, column=12).value, f"L{row}")
            target_stock = _finite_number(planning.cell(row=row, column=13).value, f"M{row}")
            debt = _finite_number(planning.cell(row=row, column=14).value, f"N{row}")
            required = _finite_number(planning.cell(row=row, column=15).value, f"O{row}")
            planned = _finite_number(planning.cell(row=row, column=16).value, f"P{row}")
            production_days = _finite_number(planning.cell(row=row, column=17).value, f"Q{row}")

            # M/P/Q không được âm. O (=p_need của engine) CÓ THỂ âm khi tồn
            # vượt nhu cầu (thừa hàng) — engine ghi p_need thô, không kẹp 0.
            for label, value in (("M", target_stock), ("P", planned), ("Q", production_days)):
                if value < -EPS:
                    raise RuntimeError(f"{label}{row} của mã {code} không được âm: {value}.")

            if per_shift <= 0 or shifts_per_day <= 0:
                if planned > EPS:
                    raise RuntimeError(
                        f"Mã {code} có P>0 nhưng E/I không hợp lệ: E={per_shift}, I={shifts_per_day}."
                    )
            else:
                # O = p_need của planning.weekly_engine (nguồn sự thật ghi workbook):
                #   - Không nợ: O = FC - tồn_sổ(K) + tồn_cuối(M).
                #   - Có nợ: theo Debt mode, KHÔNG cộng M:
                #       SUBTRACT_BOOK_ON_DEBT: O = FC + nợ - tồn_sổ(K)
                #       IGNORE_BOOK_ON_DEBT:   O = FC + nợ
                # Verifier tự lập không đọc Debt mode nên chấp nhận cả hai nhánh nợ.
                if debt > EPS:
                    candidates = (
                        forecast + debt - book_stock,
                        forecast + debt,
                    )
                    if not any(
                        math.isclose(required, cand, rel_tol=1e-9, abs_tol=1e-5)
                        for cand in candidates
                    ):
                        raise RuntimeError(
                            f"O{row} mã {code} sai: {required}; cần "
                            f"{candidates[0]} (trừ tồn sổ) hoặc {candidates[1]} (bỏ qua tồn sổ)."
                        )
                else:
                    expected_required = forecast - book_stock + target_stock
                    if not math.isclose(required, expected_required, rel_tol=1e-9, abs_tol=1e-5):
                        raise RuntimeError(
                            f"O{row} mã {code} sai: {required}; cần {expected_required}."
                        )

                # P là sản lượng CAM KẾT = min(ROUNDUP(O/mẻ hoặc /ca), đã xếp lịch).
                # Do đó P nằm trong [0, ROUNDUP(O)] chứ không nhất thiết bằng ROUNDUP(O)
                # (thiếu capacity thì P < ROUNDUP và phần thiếu là carryover). Ràng buộc
                # P = tổng SX ngày và cân bằng khối lượng được kiểm ở phần report bên dưới.
                base_qty = batch if is_sugar_classification(classification) else per_shift
                if planned > EPS and base_qty <= 0:
                    raise RuntimeError(f"Mã {code} có quantum <=0 nhưng P={planned}.")
                if base_qty > 0:
                    rounded_upper = (
                        0.0
                        if required <= EPS
                        else math.ceil(required / base_qty - 1e-12) * base_qty
                    )
                    if planned > rounded_upper + 1e-5:
                        raise RuntimeError(
                            f"P{row} mã {code}={planned} vượt ROUNDUP(O)={rounded_upper}."
                        )

                expected_q = planned / per_shift / shifts_per_day
                if not math.isclose(production_days, expected_q, rel_tol=1e-9, abs_tol=1e-6):
                    raise RuntimeError(
                        f"Q{row} mã {code} sai: {production_days}; cần {expected_q}."
                    )

            daily_values = []
            for offset, current_day in enumerate(headers):
                column = START_COLUMN + offset
                qty = _finite_number(
                    planning.cell(row=row, column=column).value,
                    f"{get_column_letter(column)}{row}",
                )
                if qty < -EPS:
                    raise RuntimeError(
                        f"Lịch mã {code} ngày {_header_day(current_day)} âm: {qty}."
                    )
                daily_values.append(qty)

            total_scheduled = sum(daily_values)
            report_balance = None
            if isinstance(schedule_report, dict):
                report_balance = (schedule_report.get("mass_balance") or {}).get(code)

            if report_balance is not None:
                report_planned = _finite_number(
                    report_balance.get("planned_qty"),
                    f"report {code}.planned_qty",
                )
                report_scheduled = _finite_number(
                    report_balance.get("scheduled_qty"),
                    f"report {code}.scheduled_qty",
                )
                carryover_qty = _finite_number(
                    report_balance.get("carryover_qty"),
                    f"report {code}.carryover_qty",
                )
                for label, value in (
                    ("planned_qty", report_planned),
                    ("scheduled_qty", report_scheduled),
                    ("carryover_qty", carryover_qty),
                ):
                    if value < -MASS_EPS:
                        raise RuntimeError(
                            f"Report mã {code} có {label} âm: {value}."
                        )
                if total_scheduled > planned + MASS_EPS:
                    raise RuntimeError(
                        f"Mã {code} scheduled={total_scheduled} vượt P={planned}; "
                        "carryover âm không được phép che sản xuất vượt kế hoạch."
                    )
                if report_scheduled > report_planned + MASS_EPS:
                    raise RuntimeError(
                        f"Report mã {code} scheduled={report_scheduled} vượt P={report_planned}."
                    )
                if carryover_qty > report_planned + MASS_EPS:
                    raise RuntimeError(
                        f"Report mã {code} carryover={carryover_qty} vượt P={report_planned}."
                    )
                if not math.isclose(report_planned, planned, rel_tol=1e-9, abs_tol=MASS_EPS):
                    raise RuntimeError(f"Report P mã {code}={report_planned} khác workbook P={planned}.")
                if not math.isclose(report_scheduled, total_scheduled, rel_tol=1e-9, abs_tol=MASS_EPS):
                    raise RuntimeError(
                        f"Report scheduled mã {code}={report_scheduled} khác workbook={total_scheduled}."
                    )
                if not math.isclose(total_scheduled + carryover_qty, planned, rel_tol=1e-9, abs_tol=MASS_EPS):
                    raise RuntimeError(
                        f"Mass balance mã {code}: scheduled {total_scheduled} + carryover {carryover_qty} != P {planned}."
                    )
            elif not math.isclose(total_scheduled, planned, rel_tol=1e-9, abs_tol=MASS_EPS):
                raise RuntimeError(
                    f"Tổng SX ngày của mã {code} = {total_scheduled} khác P={planned}. "
                    "Cần schedule report có provenance để xác nhận carryover hợp lệ."
                )

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
