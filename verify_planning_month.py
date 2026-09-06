import calendar
import math
from datetime import datetime
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import sync_planning_fc as fc
from sync_planning_calendar import build_date_headers, parse_plan_month
from sync_planning_calendar_all_months import resolve_plan_year
from sync_stock import DEST_PATH, GraphClient, get_access_token, normalize_code, to_number


PLANNING_SHEET = "Ke_hoach_SX"
STOCK_SHEET = "Ton_kho"
START_COLUMN = 19  # S


def verify_workbook(workbook_bytes):
    workbook = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        for sheet_name in (fc.FC_SHEET, STOCK_SHEET, PLANNING_SHEET):
            if sheet_name not in workbook.sheetnames:
                raise RuntimeError(f"Không tìm thấy sheet {sheet_name!r}.")

        fc_sheet = workbook[fc.FC_SHEET]
        stock_sheet = workbook[STOCK_SHEET]
        planning = workbook[PLANNING_SHEET]

        # 1) J/K phải lấy trực tiếp từ Ton_kho theo mã sản phẩm.
        stock_values = {}
        for row in range(2, stock_sheet.max_row + 1):
            code = normalize_code(stock_sheet.cell(row=row, column=1).value)
            if not code:
                continue
            if code in stock_values:
                raise RuntimeError(f"Mã {code} bị lặp trong {STOCK_SHEET}!A.")
            actual = to_number(stock_sheet.cell(row=row, column=4).value, f"{STOCK_SHEET}!D{row}")
            book = sum(
                to_number(stock_sheet.cell(row=row, column=column).value, f"{STOCK_SHEET}!{get_column_letter(column)}{row}")
                for column in range(5, 9)
            )
            stock_values[code] = (actual, book)

        stock_mismatches = []
        stock_checked = 0
        for row_number, values in enumerate(
            planning.iter_rows(min_row=2, max_col=11, values_only=True),
            start=2,
        ):
            code = normalize_code(values[0])
            if not code:
                continue
            if code not in stock_values:
                raise RuntimeError(f"Mã {code} ở {PLANNING_SHEET}!A{row_number} không có trong {STOCK_SHEET}!A.")
            expected_actual, expected_book = stock_values[code]
            actual_j = values[9]
            actual_k = values[10]
            stock_checked += 1
            if not fc._values_equal(actual_j, expected_actual) or not fc._values_equal(actual_k, expected_book):
                stock_mismatches.append((code, actual_j, expected_actual, actual_k, expected_book))

        if stock_mismatches:
            preview = "; ".join(
                f"{code}: J={actual_j!r}/Ton_kho!D={expected_j!r}, K={actual_k!r}/SUM(E:H)={expected_k!r}"
                for code, actual_j, expected_j, actual_k, expected_k in stock_mismatches[:10]
            )
            raise RuntimeError(
                f"{PLANNING_SHEET}!J:K chưa theo {STOCK_SHEET} ({len(stock_mismatches)}/{stock_checked} mã sai): {preview}"
            )

        # 2) L phải lấy đúng cột tháng được chọn tại FC!R1.
        selector = fc_sheet[fc.FC_SELECTOR_CELL].value
        selector_key = fc._normalize_header(selector)
        plan_month = parse_plan_month(selector)
        plan_year = resolve_plan_year(plan_month)

        matched_columns = []
        for column in range(fc.FC_HEADER_MIN_COL, fc.FC_HEADER_MAX_COL + 1):
            if fc._normalize_header(fc_sheet.cell(row=1, column=column).value) == selector_key:
                matched_columns.append(column)
        if len(matched_columns) != 1:
            raise RuntimeError(
                f"{fc.FC_SHEET}!{fc.FC_SELECTOR_CELL}={selector!r} phải khớp đúng 1 cột E:P, "
                f"hiện khớp {len(matched_columns)} cột."
            )
        source_column = matched_columns[0]

        fc_values = {}
        for values in fc_sheet.iter_rows(min_row=2, max_col=max(fc.FC_CODE_COL, source_column), values_only=True):
            code = normalize_code(values[fc.FC_CODE_COL - 1])
            if not code:
                continue
            fc_values[code] = values[source_column - 1]

        fc_mismatches = []
        fc_checked = 0
        for row_number, values in enumerate(
            planning.iter_rows(min_row=2, max_col=fc.PLANNING_TARGET_COL, values_only=True),
            start=2,
        ):
            code = normalize_code(values[fc.PLANNING_CODE_COL - 1])
            if not code:
                continue
            if code not in fc_values:
                raise RuntimeError(f"Mã {code} ở {PLANNING_SHEET}!A{row_number} không có trong FC!B.")
            expected = fc_values[code]
            actual = values[fc.PLANNING_TARGET_COL - 1]
            fc_checked += 1
            if not fc._values_equal(actual, expected):
                fc_mismatches.append((code, actual, expected))

        if fc_mismatches:
            preview = "; ".join(
                f"{code}: L={actual!r}, FC={expected!r}"
                for code, actual, expected in fc_mismatches[:10]
            )
            raise RuntimeError(
                f"{PLANNING_SHEET}!L chưa theo {selector!r} ({len(fc_mismatches)}/{fc_checked} mã sai): {preview}"
            )

        # 3) Header lịch phải đúng tháng và không còn cột legacy/dữ liệu dư.
        headers = build_date_headers(plan_year, plan_month)
        expected_start = headers[0]
        expected_end = headers[-1]
        end_column = START_COLUMN + len(headers) - 1
        after_end_column = end_column + 1

        actual_start = planning.cell(row=1, column=START_COLUMN).value
        actual_end = planning.cell(row=1, column=end_column).value
        if actual_start != expected_start or actual_end != expected_end:
            raise RuntimeError(
                f"Dải ngày sai cho {plan_month:02d}/{plan_year}: "
                f"S1={actual_start!r}, {get_column_letter(end_column)}1={actual_end!r}; "
                f"cần {expected_start!r} ... {expected_end!r}."
            )

        for offset, expected_header in enumerate(headers):
            column = START_COLUMN + offset
            actual_header = planning.cell(row=1, column=column).value
            if actual_header != expected_header:
                raise RuntimeError(
                    f"Tiêu đề ngày sai tại {get_column_letter(column)}1: "
                    f"{actual_header!r}; cần {expected_header!r}."
                )

        legacy_labels = {"Kỳ kế hoạch", "Tổng SX", "Chênh lệch (SX-P)"}
        for column in range(START_COLUMN, max(after_end_column + 4, 53)):
            value = planning.cell(row=1, column=column).value
            if value in legacy_labels:
                raise RuntimeError(
                    f"Còn cột legacy {value!r} tại {get_column_letter(column)}1."
                )

        stale_cells = []
        for row_number, values in enumerate(
            planning.iter_rows(
                min_row=1,
                min_col=after_end_column,
                max_col=after_end_column + 4,
                values_only=True,
            ),
            start=1,
        ):
            for offset, value in enumerate(values):
                if value not in (None, ""):
                    stale_cells.append(
                        f"{get_column_letter(after_end_column + offset)}{row_number}={value!r}"
                    )
                    if len(stale_cells) >= 10:
                        break
            if len(stale_cells) >= 10:
                break
        if stale_cells:
            raise RuntimeError(
                "Còn dữ liệu sau ngày cuối tháng: " + "; ".join(stale_cells)
            )

        # 4) Hậu kiểm độc lập dữ liệu kế hoạch và lịch ngày. Không tin line_usage
        # do allocator tự tạo: tính lại từ chính workbook cuối.
        EPS = 1e-7
        shared_lines = {"KHS", "PET 9000"}
        resource_capacity = {}
        daily_resource_usage = {}
        checked_schedule_rows = 0

        def number(value, label):
            if value in (None, ""):
                return 0.0
            if isinstance(value, bool):
                raise RuntimeError(f"{label} chứa TRUE/FALSE, không phải số.")
            try:
                result = float(value)
            except (TypeError, ValueError) as exc:
                raise RuntimeError(f"{label} không phải số: {value!r}") from exc
            if not math.isfinite(result):
                raise RuntimeError(f"{label} phải là số hữu hạn, hiện {value!r}.")
            return result

        planning_rows = []
        for row in range(2, planning.max_row + 1):
            code = normalize_code(planning.cell(row=row, column=1).value)
            if not code:
                continue

            batch = number(planning.cell(row=row, column=4).value, f"D{row}")
            per_shift = number(planning.cell(row=row, column=5).value, f"E{row}")
            line = str(planning.cell(row=row, column=6).value or "").strip()
            classification = str(planning.cell(row=row, column=8).value or "").strip()
            shifts_per_day = number(planning.cell(row=row, column=9).value, f"I{row}")
            actual_stock = number(planning.cell(row=row, column=10).value, f"J{row}")
            book_stock = number(planning.cell(row=row, column=11).value, f"K{row}")
            forecast = number(planning.cell(row=row, column=12).value, f"L{row}")
            target_stock = number(planning.cell(row=row, column=13).value, f"M{row}")
            debt = number(planning.cell(row=row, column=14).value, f"N{row}")
            required = number(planning.cell(row=row, column=15).value, f"O{row}")
            planned = number(planning.cell(row=row, column=16).value, f"P{row}")
            production_days = number(planning.cell(row=row, column=17).value, f"Q{row}")

            for label, value in (("M", target_stock), ("O", required), ("P", planned), ("Q", production_days)):
                if value < -EPS:
                    raise RuntimeError(f"{label}{row} của mã {code} không được âm: {value}.")

            if per_shift <= 0 or shifts_per_day <= 0:
                if planned > EPS:
                    raise RuntimeError(
                        f"Mã {code} có P>0 nhưng E/I không hợp lệ: E={per_shift}, I={shifts_per_day}."
                    )
            else:
                urgent_qty = max(forecast + debt - max(actual_stock, 0.0), 0.0)
                if urgent_qty > EPS:
                    expected_required = urgent_qty
                else:
                    planning_stock = min(max(actual_stock, 0.0), max(book_stock, 0.0))
                    expected_required = max(
                        forecast + debt + target_stock - planning_stock,
                        0.0,
                    )
                if not math.isclose(required, expected_required, rel_tol=1e-9, abs_tol=1e-5):
                    raise RuntimeError(
                        f"O{row} mã {code} sai: {required}; cần {expected_required}."
                    )

                base_qty = batch if classification.casefold() == "có đường".casefold() else per_shift
                if planned > EPS and base_qty <= 0:
                    raise RuntimeError(f"Mã {code} có quantum <=0 nhưng P={planned}.")
                expected_planned = (
                    0.0
                    if expected_required <= EPS
                    else math.ceil(expected_required / base_qty - 1e-12) * base_qty
                )
                if not math.isclose(planned, expected_planned, rel_tol=1e-9, abs_tol=1e-5):
                    raise RuntimeError(
                        f"P{row} mã {code} sai: {planned}; cần {expected_planned}."
                    )

                expected_q = planned / per_shift / shifts_per_day
                if not math.isclose(production_days, expected_q, rel_tol=1e-9, abs_tol=1e-6):
                    raise RuntimeError(
                        f"Q{row} mã {code} sai: {production_days}; cần {expected_q}."
                    )

            daily_values = []
            for offset, current_day in enumerate(headers):
                column = START_COLUMN + offset
                qty = number(
                    planning.cell(row=row, column=column).value,
                    f"{get_column_letter(column)}{row}",
                )
                if qty < -EPS:
                    raise RuntimeError(
                        f"Lịch mã {code} ngày {current_day:%d/%m} âm: {qty}."
                    )
                daily_values.append(qty)

            total_scheduled = sum(daily_values)
            if not math.isclose(total_scheduled, planned, rel_tol=1e-9, abs_tol=1e-5):
                raise RuntimeError(
                    f"Tổng SX ngày của mã {code} = {total_scheduled} khác P={planned}. "
                    "Nếu thiếu capacity phải báo carryover/infeasible, không để workbook im lặng lệch mass balance."
                )

            resource = "KHS + PET 9000" if line in shared_lines else line
            if resource:
                existing_capacity = resource_capacity.get(resource)
                if existing_capacity is None:
                    resource_capacity[resource] = shifts_per_day
                elif resource == "KHS + PET 9000" or line not in {"RGB", "Galon"}:
                    resource_capacity[resource] = min(existing_capacity, shifts_per_day)
                else:
                    resource_capacity[resource] = max(existing_capacity, shifts_per_day)

                if per_shift > EPS:
                    for current_day, qty in zip(headers, daily_values):
                        key = (resource, current_day)
                        daily_resource_usage[key] = daily_resource_usage.get(key, 0.0) + qty / per_shift

            planning_rows.append(code)
            checked_schedule_rows += 1

        for (resource, current_day), used_shifts in daily_resource_usage.items():
            capacity = resource_capacity.get(resource, 0.0)
            if used_shifts > capacity + 1e-6:
                raise RuntimeError(
                    f"Resource {resource} ngày {current_day:%d/%m} có ít nhất "
                    f"{used_shifts:.3f} ca sản xuất > capacity {capacity:.3f}; "
                    "chưa tính setup."
                )

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
    verify_workbook(dest_bytes)


if __name__ == "__main__":
    main()
