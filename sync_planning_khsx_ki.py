"""Module tổng hợp kế hoạch sản xuất theo tuần cho sheet KHSX_ki.

Quy tắc nghiệp vụ:
- Sheet Ke_hoach_SX là nguồn kế hoạch sản xuất chi tiết đã được engine phân bổ theo ngày
  và kiểm tra công suất.
- Sheet KHSX_ki là kết quả tổng hợp phân bổ theo tuần từ lịch ngày đó.
- Không chạy thuật toán lập lịch độc lập cho KHSX_ki.
- Chu kỳ tuần theo chuẩn Monday-Sunday (Thứ Hai đến Chủ Nhật).
  Tuần 1 bắt đầu từ ngày 01 đến Chủ Nhật đầu tiên của tháng.
  Các tuần tiếp theo từ Thứ Hai đến Chủ Nhật (hoặc cuối tháng).
- Tổng sản lượng các tuần cho mỗi mã sản phẩm (ma_sp) khớp 100% với tổng ngày
  và bằng sản lượng cam kết cột P của tháng đầy đủ trong Ke_hoach_SX.
"""
from __future__ import annotations

import calendar
import datetime
import math
import re
from collections import defaultdict
from io import BytesIO
from typing import Any

from openpyxl import load_workbook

import sync_stock

KHSX_KI_SHEET = "KHSX_ki"
PLANNING_SHEET = "Ke_hoach_SX"
START_DAILY_COLUMN = 19  # Cột S trong Ke_hoach_SX
P_COLUMN = 16            # Cột P trong Ke_hoach_SX
EPS = 1e-6


def has_khsx_ki_sheet(workbook_bytes: bytes) -> bool:
    """Kiểm tra sự tồn tại của sheet KHSX_ki trong workbook mà không nạp toàn bộ dữ liệu."""
    wb = load_workbook(BytesIO(workbook_bytes), read_only=True)
    try:
        return KHSX_KI_SHEET in wb.sheetnames
    finally:
        wb.close()


def compute_month_weeks(
    year: int,
    month: int,
    *,
    num_week_cols: int = 5,
) -> list[dict[str, Any]]:
    """Phân bổ các ngày trong tháng thành các tuần theo chuẩn Monday-Sunday.

    Tuần 1: từ ngày 1 đến Chủ Nhật đầu tiên.
    Các tuần tiếp theo: từ Thứ Hai đến Chủ Nhật.
    Tuần cuối cùng: kết thúc vào ngày cuối cùng của tháng (days_in_month).
    Đảm bảo phủ 100% các ngày trong tháng không trùng lặp và không có khoảng trống.
    """
    days_in_month = calendar.monthrange(year, month)[1]
    weeks: list[dict[str, Any]] = []
    w_start = 1

    for w_idx in range(1, num_week_cols + 1):
        if w_start > days_in_month:
            weeks.append({
                "week_num": w_idx,
                "start_day": 0,
                "end_day": 0,
                "days": [],
                "label": f"Tuần {w_idx}\n-",
            })
            continue

        if w_idx == num_week_cols:
            w_end = days_in_month
        else:
            curr = w_start
            while curr < days_in_month and datetime.date(year, month, curr).weekday() != 6:
                curr += 1
            w_end = curr

        days = list(range(w_start, w_end + 1))
        label = f"Tuần {w_idx}\n{w_start:02d}/{month:02d}-{w_end:02d}/{month:02d}"
        weeks.append({
            "week_num": w_idx,
            "start_day": w_start,
            "end_day": w_end,
            "days": days,
            "label": label,
        })
        w_start = w_end + 1

    return weeks


def parse_week_columns(
    ws,
    *,
    plan_year: int,
    plan_month: int,
    header_row: int = 6,
) -> tuple[list[dict[str, Any]], int]:
    """Xác định danh sách các cột tuần và cột Tổng cộng từ dòng tiêu đề của KHSX_ki.

    Trả về (weeks, total_col).
    Nếu tiêu đề đã có ngày khớp kỳ kế hoạch, sử dụng dải ngày đó;
    ngược lại tự động sinh dải ngày chuẩn Monday-Sunday và cập nhật tiêu đề.
    """
    days_in_month = calendar.monthrange(plan_year, plan_month)[1]
    total_col: int | None = None
    week_cols: list[int] = []

    for col in range(5, ws.max_column + 1):
        val = str(ws.cell(header_row, col).value or "").strip()
        if not val:
            continue
        val_norm = val.casefold()
        if "tổng cộng" in val_norm or "tong cong" in val_norm:
            total_col = col
            break
        if "tuần" in val_norm or "tuan" in val_norm or re.search(r"\d{1,2}/\d{1,2}", val):
            week_cols.append(col)

    if total_col is None:
        # Mặc định cột J (10) nếu chưa gắn nhãn
        total_col = max(week_cols) + 1 if week_cols else 10
        if not week_cols:
            week_cols = list(range(5, total_col))

    num_weeks = len(week_cols)
    std_weeks = compute_month_weeks(plan_year, plan_month, num_week_cols=num_weeks)

    weeks: list[dict[str, Any]] = []
    for idx, col in enumerate(week_cols):
        cell_val = str(ws.cell(header_row, col).value or "").strip()
        m = re.search(r"(\d{1,2})/(\d{1,2})\s*-\s*(\d{1,2})/(\d{1,2})", cell_val)
        matched = False
        if m:
            s_d, s_m, e_d, e_m = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
            if s_m == plan_month and e_m == plan_month and 1 <= s_d <= e_d <= days_in_month:
                weeks.append({
                    "col": col,
                    "week_num": idx + 1,
                    "start_day": s_d,
                    "end_day": e_d,
                    "days": list(range(s_d, e_d + 1)),
                    "label": cell_val,
                })
                matched = True

        if not matched:
            w_info = dict(std_weeks[idx])
            w_info["col"] = col
            weeks.append(w_info)
            # Cập nhật tiêu đề hiển thị đúng kỳ
            ws.cell(header_row, col).value = w_info["label"]

    return weeks, total_col


def _read_daily_plan_from_analysis_or_sheet(
    workbook_bytes: bytes,
    weekly_analysis: Any | None,
    *,
    plan_year: int,
    plan_month: int,
) -> tuple[dict[tuple[int, int], float], dict[int, float]]:
    """Trích xuất kế hoạch sản xuất theo ngày và sản lượng cam kết cột P theo mã SP.

    Trả về (daily_map, committed_map):
    - daily_map[(code, day)] = qty
    - committed_map[code] = committed_qty
    """
    daily_map: dict[tuple[int, int], float] = defaultdict(float)
    committed_map: dict[int, float] = {}

    if weekly_analysis is not None and getattr(weekly_analysis, "daily_plan", None):
        for item in weekly_analysis.daily_plan:
            code = int(item.ma_sp)
            day = item.date.day
            daily_map[(code, day)] += float(item.qty)

        scheduled_by_code: dict[int, float] = defaultdict(float)
        for item in weekly_analysis.daily_plan:
            scheduled_by_code[int(item.ma_sp)] += float(item.qty)

        for calc in weekly_analysis.calculated:
            code = int(calc.input.ma_sp)
            scheduled = max(0.0, scheduled_by_code.get(code, 0.0))
            committed_map[code] = min(calc.schedulable_qty, scheduled)
        return daily_map, committed_map

    # Fallback: đọc trực tiếp từ sheet Ke_hoach_SX trong workbook_bytes
    wb = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        ws = wb[PLANNING_SHEET]
        days_in_month = calendar.monthrange(plan_year, plan_month)[1]
        for r in range(2, ws.max_row + 1):
            raw_code = ws.cell(r, 1).value
            code = sync_stock.normalize_code(raw_code)
            if not code:
                continue
            code_int = int(code)
            committed_map[code_int] = float(ws.cell(r, P_COLUMN).value or 0.0)
            for d in range(1, days_in_month + 1):
                col = START_DAILY_COLUMN + d - 1
                qty = float(ws.cell(r, col).value or 0.0)
                if qty > EPS:
                    daily_map[(code_int, d)] += qty
        return daily_map, committed_map
    finally:
        wb.close()


def patch_khsx_ki_workbook(
    workbook_bytes: bytes,
    *,
    weekly_analysis: Any | None = None,
    plan_year: int,
    plan_month: int,
) -> tuple[bytes, dict[str, Any]]:
    """Tổng hợp lịch ngày từ Ke_hoach_SX vào sheet KHSX_ki, bảo toàn định dạng."""
    wb = load_workbook(BytesIO(workbook_bytes), data_only=False)
    try:
        if KHSX_KI_SHEET not in wb.sheetnames:
            return workbook_bytes, {"ok": False, "reason": f"Thiếu sheet {KHSX_KI_SHEET}"}

        ws = wb[KHSX_KI_SHEET]
        weeks, total_col = parse_week_columns(ws, plan_year=plan_year, plan_month=plan_month)
        daily_map, committed_map = _read_daily_plan_from_analysis_or_sheet(
            workbook_bytes,
            weekly_analysis,
            plan_year=plan_year,
            plan_month=plan_month,
        )

        # Cập nhật ô tiêu đề kỳ và ngày lập nếu có
        period_text = f"Kỳ kế hoạch: Tháng {plan_month}"
        date_text = f"Ngày lập: 01/{plan_month:02d}/{plan_year}"
        if ws.cell(3, 1).value is not None:
            ws.cell(3, 1).value = period_text
        if ws.cell(4, 1).value is not None:
            ws.cell(4, 1).value = date_text

        # Tìm các dòng sản phẩm và dòng TỔNG CỘNG
        total_row: int | None = None
        data_rows: list[tuple[int, int]] = []  # (row_idx, code_int)

        for r in range(7, ws.max_row + 1):
            raw_a = str(ws.cell(r, 1).value or "").strip().casefold()
            if "tổng cộng" in raw_a or "tong cong" in raw_a:
                total_row = r
                break
            raw_code = ws.cell(r, 2).value
            code = sync_stock.normalize_code(raw_code)
            if code:
                data_rows.append((r, int(code)))

        col_totals: dict[int, float] = defaultdict(float)
        sku_reports = []

        # Ghi dữ liệu từng dòng SKU
        for r, code in data_rows:
            row_total = 0.0
            sku_week_vals = {}
            for w in weeks:
                col = w["col"]
                w_qty = sum(daily_map.get((code, d), 0.0) for d in w["days"])
                val = float(w_qty) if w_qty > EPS else 0.0
                ws.cell(r, col).value = val
                col_totals[col] += val
                row_total += val
                sku_week_vals[f"week_{w['week_num']}"] = val

            ws.cell(r, total_col).value = float(row_total) if row_total > EPS else 0.0
            col_totals[total_col] += row_total
            sku_reports.append({
                "code": code,
                "weeks": sku_week_vals,
                "total": row_total,
                "committed_p": committed_map.get(code, 0.0),
            })

        # Ghi dòng TỔNG CỘNG nếu có
        if total_row is not None:
            for w in weeks:
                col = w["col"]
                ws.cell(total_row, col).value = float(col_totals[col])
            ws.cell(total_row, total_col).value = float(col_totals[total_col])

        out = BytesIO()
        wb.save(out)
        patched_bytes = out.getvalue()

        report = {
            "ok": True,
            "sheet": KHSX_KI_SHEET,
            "period": f"{plan_year:04d}-{plan_month:02d}",
            "checked_skus": len(data_rows),
            "week_columns": [w["label"].replace("\n", " ") for w in weeks],
            "total_production": col_totals[total_col],
            "grand_total": col_totals[total_col],
        }
        return patched_bytes, report
    finally:
        wb.close()


def verify_khsx_ki(
    workbook_bytes: bytes,
    *,
    plan_year: int,
    plan_month: int,
) -> dict[str, Any]:
    """Kiểm tra độc lập tính toàn vẹn số liệu của sheet KHSX_ki.

    Các bất biến bắt buộc:
    1. Sheet KHSX_ki tồn tại.
    2. Toàn bộ ô sản lượng tuần, tổng dòng, tổng cột là số thực hữu hạn >= 0.
    3. Với từng SKU: Tổng các tuần == Cột J (Tổng cộng) (dung sai < 1e-6).
    4. Cột J == Tổng lịch ngày trong Ke_hoach_SX == Cột P (dung sai < 1e-6).
    5. Dòng TỔNG CỘNG khớp tổng các dòng của từng cột tuần.
    6. Ô grand total khớp tổng các dòng và tổng các cột.
    """
    wb = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        if KHSX_KI_SHEET not in wb.sheetnames:
            return {"ok": True, "present": False, "message": "Sheet KHSX_ki không có trong workbook"}

        ws_ki = wb[KHSX_KI_SHEET]
        if PLANNING_SHEET not in wb.sheetnames:
            raise RuntimeError(f"Không tìm thấy sheet {PLANNING_SHEET} để đối chiếu KHSX_ki.")
        ws_kh = wb[PLANNING_SHEET]

        days_in_month = calendar.monthrange(plan_year, plan_month)[1]
        weeks, total_col = parse_week_columns(ws_ki, plan_year=plan_year, plan_month=plan_month)

        # Đọc lịch ngày và cột P từ Ke_hoach_SX
        kh_daily: dict[int, list[float]] = {}
        kh_p: dict[int, float] = {}
        for r in range(2, ws_kh.max_row + 1):
            raw = ws_kh.cell(r, 1).value
            code = sync_stock.normalize_code(raw)
            if not code:
                continue
            code_int = int(code)
            kh_p[code_int] = float(ws_kh.cell(r, P_COLUMN).value or 0.0)
            kh_daily[code_int] = [
                float(ws_kh.cell(r, START_DAILY_COLUMN + d - 1).value or 0.0)
                for d in range(1, days_in_month + 1)
            ]

        # Đọc KHSX_ki
        total_row: int | None = None
        ki_data_rows: list[tuple[int, int]] = []
        for r in range(7, ws_ki.max_row + 1):
            raw_a = str(ws_ki.cell(r, 1).value or "").strip().casefold()
            if "tổng cộng" in raw_a or "tong cong" in raw_a:
                total_row = r
                break
            raw_code = ws_ki.cell(r, 2).value
            code = sync_stock.normalize_code(raw_code)
            if code:
                ki_data_rows.append((r, int(code)))

        col_sums: dict[int, float] = defaultdict(float)
        checked_count = 0

        for r, code in ki_data_rows:
            week_vals = []
            for w in weeks:
                val = ws_ki.cell(r, w["col"]).value
                if val is None:
                    val = 0.0
                if not isinstance(val, (int, float)) or not math.isfinite(val) or val < -EPS:
                    raise RuntimeError(f"KHSX_ki!R{r}C{w['col']} có giá trị không hợp lệ: {val!r}")
                week_vals.append(float(val))
                col_sums[w["col"]] += float(val)

            row_total = ws_ki.cell(r, total_col).value
            if row_total is None:
                row_total = 0.0
            if not isinstance(row_total, (int, float)) or not math.isfinite(row_total) or row_total < -EPS:
                raise RuntimeError(f"KHSX_ki!R{r}C{total_col} (Tổng cộng) không hợp lệ: {row_total!r}")
            row_total = float(row_total)
            col_sums[total_col] += row_total

            # 1. Tổng các tuần phải bằng cột J
            sum_weeks = sum(week_vals)
            if abs(sum_weeks - row_total) > EPS:
                raise RuntimeError(
                    f"KHSX_ki mã {code} (dòng {r}): Tổng tuần ({sum_weeks}) != Cột tổng ({row_total})"
                )

            # 2. Đối chiếu với Ke_hoach_SX nếu mã có trong lịch sản xuất
            if code in kh_daily:
                daily_sum = sum(kh_daily[code])
                p_qty = kh_p.get(code, 0.0)

                if abs(row_total - daily_sum) > EPS:
                    raise RuntimeError(
                        f"KHSX_ki mã {code}: Tổng tuần ({row_total}) != Tổng lịch ngày ({daily_sum})"
                    )
                if abs(row_total - p_qty) > EPS:
                    raise RuntimeError(
                        f"KHSX_ki mã {code}: Tổng tuần ({row_total}) != Cột P cam kết ({p_qty})"
                    )

                # Đối chiếu từng tuần với tổng ngày tương ứng
                for idx, w in enumerate(weeks):
                    expected_w = sum(kh_daily[code][d - 1] for d in w["days"])
                    actual_w = week_vals[idx]
                    if abs(actual_w - expected_w) > EPS:
                        raise RuntimeError(
                            f"KHSX_ki mã {code} Tuần {w['week_num']}: Thực tế {actual_w} != Kỳ vọng {expected_w}"
                        )

            checked_count += 1

        # Kiểm tra dòng TỔNG CỘNG
        if total_row is not None:
            for w in weeks:
                col = w["col"]
                actual_col_tot = float(ws_ki.cell(total_row, col).value or 0.0)
                expected_col_tot = col_sums[col]
                if abs(actual_col_tot - expected_col_tot) > EPS:
                    raise RuntimeError(
                        f"KHSX_ki dòng TỔNG CỘNG cột {col}: {actual_col_tot} != Tổng dòng ({expected_col_tot})"
                    )

            actual_grand = float(ws_ki.cell(total_row, total_col).value or 0.0)
            expected_grand = col_sums[total_col]
            if abs(actual_grand - expected_grand) > EPS:
                raise RuntimeError(
                    f"KHSX_ki grand total: {actual_grand} != Tổng cột {expected_grand}"
                )

        return {
            "ok": True,
            "present": True,
            "checked_skus": checked_count,
            "grand_total": col_sums[total_col],
            "period": f"{plan_year:04d}-{plan_month:02d}",
        }
    finally:
        wb.close()
