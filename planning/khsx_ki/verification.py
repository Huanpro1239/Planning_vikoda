"""Independent read-only verification for KHSX_ki."""

from __future__ import annotations

import calendar
import math
from collections import defaultdict
from io import BytesIO
from typing import Any

from openpyxl import load_workbook

from .calendar import compute_standard_calendar_weeks
from .layout import get_layout_spec
from .workbook import (
    EPS,
    KHSX_KI_SHEET,
    PLANNING_SHEET,
    _read_khsx_ki_skus,
    _read_planning_sheet_data,
    reconcile_skus,
)


def verify_khsx_ki(
    workbook_bytes: bytes,
    *,
    plan_year: int,
    plan_month: int,
) -> dict[str, Any]:
    """Kiểm tra độc lập tính toàn vẹn số liệu và cấu trúc của sheet KHSX_ki.

    Các bất biến bắt buộc:
    1. Sheet KHSX_ki tồn tại.
    2. Toàn bộ ô sản lượng tuần, tổng dòng, tổng cột là số thực hữu hạn >= 0.
    3. Phân hoạch tuần khớp 100% lịch Monday-Sunday (không gap, không overlap).
    4. Tiêu đề dòng 6 khớp chính xác phân hoạch chuẩn theo lịch.
    5. Tập mã SKU duy nhất và khớp 100% giữa Ke_hoach_SX và KHSX_ki (chặn trùng, thiếu, thừa).
    6. Với từng SKU: Tổng các tuần == Cột Tổng cộng == Tổng lịch ngày Ke_hoach_SX == Cột P.
    7. Dòng TỔNG CỘNG khớp tổng các dòng của từng cột tuần.
    8. Ô grand total khớp tổng các dòng và tổng các cột.
    9. Với tháng <= 5 tuần, cột 11 (K) trong vùng quản lý không được chứa dữ liệu dư thừa.
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
        std_weeks = compute_standard_calendar_weeks(plan_year, plan_month)
        weeks, total_col = get_layout_spec(plan_year, plan_month)

        # 1. Xác minh phân hoạch tuần phủ 100% các ngày trong tháng không trùng/khuyết
        active_days = [d for w in weeks for d in w["days"]]
        if sorted(active_days) != list(range(1, days_in_month + 1)):
            raise RuntimeError(
                f"Phân hoạch tuần KHSX_ki không phủ đúng dải ngày [1, {days_in_month}] của tháng {plan_month}."
            )
        if len(active_days) != len(set(active_days)):
            raise RuntimeError("Phân hoạch tuần KHSX_ki có ngày bị chồng lấn (trùng lặp).")

        # 2. Kiểm tra tiêu đề dòng 6 (Header)
        for w in weeks:
            col = w["col"]
            val = str(ws_ki.cell(6, col).value or "").strip()
            if w["days"]:
                expected_range = f"{w['start_day']:02d}/{plan_month:02d}-{w['end_day']:02d}/{plan_month:02d}"
                clean_val = val.replace(" ", "")
                if expected_range not in clean_val:
                    raise RuntimeError(
                        f"KHSX_ki tiêu đề cột {col} ({val!r}) không khớp dải ngày chuẩn {expected_range}."
                    )
                if f"Tuần{w['week_num']}" not in clean_val and f"tuần{w['week_num']}" not in clean_val.lower():
                    raise RuntimeError(
                        f"KHSX_ki tiêu đề cột {col} ({val!r}) không đúng tên Tuần {w['week_num']}."
                    )
            else:
                if "5" not in val or ("-" not in val and val != ""):
                    raise RuntimeError(
                        f"KHSX_ki tiêu đề cột {col} ({val!r}) không đúng định dạng tuần rỗng ('Tuần 5\n-')."
                    )

        val_total = str(ws_ki.cell(6, total_col).value or "").strip().casefold()
        if "tổng cộng" not in val_total and "tong cong" not in val_total:
            raise RuntimeError(
                f"KHSX_ki tiêu đề cột {total_col} ({val_total!r}) không phải cột Tổng cộng."
            )

        if len(std_weeks) <= 5:
            val_k = ws_ki.cell(6, 11).value
            if val_k is not None and str(val_k).strip() != "":
                raise RuntimeError(
                    f"KHSX_ki tiêu đề cột 11 (K) vẫn còn dữ liệu ({val_k!r}) dù tháng {plan_month} chỉ có {len(std_weeks)} tuần."
                )

        # 3. Đọc dữ liệu Ke_hoach_SX và KHSX_ki, đối chiếu 100% SKU
        seen_kh, kh_p, kh_daily = _read_planning_sheet_data(
            ws_kh,
            plan_year=plan_year,
            plan_month=plan_month,
        )
        seen_ki, total_row = _read_khsx_ki_skus(ws_ki, header_row=6)
        reconcile_skus(seen_kh, seen_ki)

        col_sums: dict[int, float] = defaultdict(float)
        checked_count = 0

        # 4. Kiểm tra từng dòng SKU
        for code, r in seen_ki.items():
            week_vals = []
            for w in weeks:
                val = ws_ki.cell(r, w["col"]).value
                if val is None:
                    val = 0.0
                if not isinstance(val, (int, float)) or not math.isfinite(val) or val < -EPS:
                    raise RuntimeError(f"KHSX_ki!R{r}C{w['col']} có giá trị không hợp lệ: {val!r}")
                val = float(val)
                if not w["days"] and abs(val) > EPS:
                    raise RuntimeError(f"KHSX_ki!R{r}C{w['col']} thuộc tuần không có ngày nhưng có giá trị: {val}")
                week_vals.append(val)
                col_sums[w["col"]] += val

            row_total = ws_ki.cell(r, total_col).value
            if row_total is None:
                row_total = 0.0
            if not isinstance(row_total, (int, float)) or not math.isfinite(row_total) or row_total < -EPS:
                raise RuntimeError(f"KHSX_ki!R{r}C{total_col} (Tổng cộng) không hợp lệ: {row_total!r}")
            row_total = float(row_total)
            col_sums[total_col] += row_total

            # Đối chiếu 1: Tổng các tuần phải bằng cột Tổng cộng
            sum_weeks = sum(week_vals)
            if abs(sum_weeks - row_total) > EPS:
                raise RuntimeError(
                    f"KHSX_ki mã {code} (dòng {r}): Tổng tuần ({sum_weeks}) != Cột tổng ({row_total})"
                )

            # Đối chiếu 2: Cột Tổng cộng phải khớp tổng lịch ngày trong Ke_hoach_SX
            daily_sum = sum(kh_daily.get((code, d), 0.0) for d in range(1, days_in_month + 1))
            if abs(row_total - daily_sum) > EPS:
                raise RuntimeError(
                    f"KHSX_ki mã {code} (dòng {r}): Cột tổng ({row_total}) != Tổng lịch ngày ({daily_sum})"
                )

            # Đối chiếu 3: Cột Tổng cộng phải khớp cột P cam kết trong Ke_hoach_SX
            p_qty = kh_p.get(code, 0.0)
            if abs(row_total - p_qty) > EPS:
                raise RuntimeError(
                    f"KHSX_ki mã {code} (dòng {r}): Cột tổng ({row_total}) != Cột P cam kết ({p_qty})"
                )

            # Đối chiếu 4: Từng tuần phải khớp chính xác tổng các ngày tương ứng trong lịch
            for idx, w in enumerate(weeks):
                expected_w = sum(kh_daily.get((code, d), 0.0) for d in w["days"])
                actual_w = week_vals[idx]
                if abs(actual_w - expected_w) > EPS:
                    raise RuntimeError(
                        f"KHSX_ki mã {code} Tuần {w['week_num']} (dòng {r}): "
                        f"Thực tế {actual_w} != Kỳ vọng {expected_w}"
                    )

            if len(std_weeks) <= 5:
                val_k = ws_ki.cell(r, 11).value
                if val_k is not None and str(val_k).strip() != "":
                    raise RuntimeError(f"KHSX_ki!R{r}C11 vẫn có dữ liệu ({val_k!r}) dù tháng chỉ có {len(std_weeks)} tuần.")

            checked_count += 1

        # 5. Kiểm tra dòng TỔNG CỘNG
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

            if len(std_weeks) <= 5:
                val_k = ws_ki.cell(total_row, 11).value
                if val_k is not None and str(val_k).strip() != "":
                    raise RuntimeError(f"KHSX_ki!R{total_row}C11 vẫn có dữ liệu ({val_k!r}) dù tháng chỉ có {len(std_weeks)} tuần.")

        return {
            "ok": True,
            "present": True,
            "checked_skus": checked_count,
            "grand_total": col_sums[total_col],
            "period": f"{plan_year:04d}-{plan_month:02d}",
            "num_weeks": len(std_weeks),
        }
    finally:
        wb.close()

