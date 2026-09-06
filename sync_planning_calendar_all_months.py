import zipfile
from datetime import datetime
from io import BytesIO

import sync_planning_calendar as calendar_sync


def resolve_plan_year(plan_month, now=None):
    now = now or datetime.now(calendar_sync.TIMEZONE)
    year = now.year
    delta = plan_month - now.month

    # Chọn năm gần nhất quanh thời điểm hiện tại để xử lý Dec -> Jan và Jan -> Dec.
    if delta <= -6:
        year += 1
    elif delta >= 6:
        year -= 1
    return year


def prepare_calendar_update_all_months(workbook_bytes, *, plan_year=None):
    if plan_year is None:
        with zipfile.ZipFile(BytesIO(workbook_bytes), "r") as archive:
            selector = calendar_sync._read_selector_from_archive(archive)
        plan_month = calendar_sync.parse_plan_month(selector)
        plan_year = resolve_plan_year(plan_month)

    return _ORIGINAL_PREPARE(workbook_bytes, plan_year=plan_year)


_ORIGINAL_PREPARE = calendar_sync.prepare_calendar_update
calendar_sync.prepare_calendar_update = prepare_calendar_update_all_months


if __name__ == "__main__":
    calendar_sync.main_with_retry()
