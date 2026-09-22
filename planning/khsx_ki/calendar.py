"""Calendar partitioning for KHSX_ki."""

from __future__ import annotations

import calendar
import datetime
from typing import Any


def compute_standard_calendar_weeks(
    year: int,
    month: int,
) -> list[dict[str, Any]]:
    """Phân bổ các ngày trong tháng thành các tuần theo chuẩn Monday-Sunday.

    Quy tắc:
    - Tuần 1: từ ngày 1 đến Chủ Nhật đầu tiên của tháng (hoặc ngày cuối tháng nếu tháng kết thúc trước).
    - Các tuần tiếp theo: từ Thứ Hai đến Chủ Nhật (hoặc ngày cuối tháng nếu tháng kết thúc trước).
    - Độ dài mỗi tuần từ 1 đến 7 ngày. Không tuần nào dài quá 7 ngày.
    - Không tuần nào vượt ra ngoài tháng (1 <= start_day <= end_day <= days_in_month).
    - Tập hợp các ngày của tất cả các tuần là phân hoạch chính xác của [1, days_in_month]
      (không có ngày trùng lặp, không có ngày bị bỏ sót).
    Số tuần W sẽ là 4, 5, hoặc 6 tùy theo tháng.
    """
    days_in_month = calendar.monthrange(year, month)[1]
    weeks: list[dict[str, Any]] = []
    w_start = 1
    w_idx = 1
    while w_start <= days_in_month:
        curr = w_start
        # Monday is 0, Sunday is 6
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
        w_idx += 1
    return weeks

def compute_month_weeks(
    year: int,
    month: int,
    *,
    num_week_cols: int | None = None,
) -> list[dict[str, Any]]:
    """Trả về danh sách tuần theo chuẩn Monday-Sunday.

    Nếu num_week_cols được chỉ định và lớn hơn số tuần lịch thực tế (ví dụ tháng 4 tuần
    nhưng template có 5 cột), các tuần bổ sung sẽ là tuần trống (start_day=0, end_day=0, days=[]).
    Nếu số tuần lịch thực tế lớn hơn num_week_cols (ví dụ tháng 6 tuần nhưng num_week_cols=5),
    vẫn trả về đầy đủ các tuần thực tế để không bị gộp/ép ngày vi phạm ranh giới tuần.
    """
    std_weeks = compute_standard_calendar_weeks(year, month)
    if num_week_cols is None or num_week_cols <= len(std_weeks):
        return std_weeks

    # Padding thêm tuần rỗng nếu num_week_cols > len(std_weeks) (ví dụ tháng 4 tuần trong template 5 cột)
    padded = list(std_weeks)
    for w_idx in range(len(std_weeks) + 1, num_week_cols + 1):
        padded.append({
            "week_num": w_idx,
            "start_day": 0,
            "end_day": 0,
            "days": [],
            "label": f"Tuần {w_idx}\n-",
        })
    return padded

