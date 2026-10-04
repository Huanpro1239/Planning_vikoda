"""Weekly schedule workbook patching."""

from __future__ import annotations

import calendar
from collections import defaultdict
from datetime import date
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

from .inputs import EPS, PLANNING_SHEET, START_COLUMN
from .schedule import (
    WeeklyAnalysis,
    committed_days,
    committed_qty,
    scheduled_by_code,
)


def _write_shortfall_notes(workbook, analysis, committed):
    name = "Ghi_chu_Planning"
    if name in workbook.sheetnames:
        del workbook[name]
    notes = workbook.create_sheet(name)
    notes.append(["Kỳ kế hoạch", "Mã sản phẩm", "Tên sản phẩm", "ĐVT",
                  "Mục tiêu SX", "Đã xếp SX", "Thiếu bán hàng/nợ",
                  "Thiếu tồn an toàn", "Ghi chú"])
    for calc in analysis.calculated:
        done = committed_qty(calc, committed)
        service_missing = max(0.0, calc.service_qty - done)
        buffer_missing = max(0.0, calc.buffer_qty - max(0.0, done - calc.service_qty))
        if service_missing <= 1e-5 and buffer_missing <= 1e-5:
            continue
        parts = []
        if service_missing > 1e-5:
            parts.append(f"Thiếu bán hàng/nợ: {service_missing:,.2f} {calc.input.don_vi_tinh}")
        if buffer_missing > 1e-5:
            parts.append(f"Thiếu tồn an toàn: {buffer_missing:,.2f} {calc.input.don_vi_tinh}")
        notes.append([f"{analysis.period_year}-{analysis.period_month:02d}",
                      str(calc.input.ma_sp), calc.input.ten_sp, calc.input.don_vi_tinh,
                      calc.schedulable_qty, done, service_missing, buffer_missing,
                      "; ".join(parts) + ". Vẫn chạy theo lịch đã xếp; phần thiếu chưa được xếp SX."])
    notes.freeze_panes = "A2"
    notes.auto_filter.ref = notes.dimensions
    for cell in notes[1]:
        cell.font = Font(bold=True)
    for column in "ABCDEFGH":
        notes.column_dimensions[column].width = 22
    notes.column_dimensions["C"].width = 42
    notes.column_dimensions["I"].width = 85
    for cells in notes.iter_rows(min_row=2):
        cells[8].alignment = Alignment(wrap_text=True, vertical="top")
        notes.row_dimensions[cells[0].row].height = 45
        for cell in cells[4:8]:
            cell.number_format = "#,##0.00"


def patch_weekly_workbook(
    workbook_bytes: bytes,
    analysis: WeeklyAnalysis,
) -> bytes:
    workbook = load_workbook(
        BytesIO(workbook_bytes),
        data_only=False,
    )
    try:
        worksheet = workbook[PLANNING_SHEET]
        lookup: dict[tuple[int, date], float] = defaultdict(float)
        for item in analysis.daily_plan:
            lookup[
                (item.ma_sp, item.date)
            ] += float(item.qty)

        days = calendar.monthrange(
            analysis.period_year,
            analysis.period_month,
        )[1]
        committed = scheduled_by_code(
            analysis.daily_plan
        )

        for calc in analysis.calculated:
            row = calc.input.source_row
            committed_value = committed_qty(
                calc,
                committed,
            )
            committed_day_value = committed_days(
                calc,
                committed,
            )

            worksheet.cell(
                row,
                15,
                calc.p_need,
            )
            worksheet.cell(
                row,
                16,
                committed_value,
            )
            worksheet.cell(
                row,
                17,
                committed_day_value,
            )
            worksheet.cell(
                row,
                18,
                calc.start_datetime,
            )

            for day in range(1, days + 1):
                qty = lookup.get(
                    (
                        calc.input.ma_sp,
                        date(
                            analysis.period_year,
                            analysis.period_month,
                            day,
                        ),
                    ),
                    0.0,
                )
                worksheet.cell(
                    row,
                    START_COLUMN + day - 1,
                ).value = (
                    qty
                    if qty > EPS
                    else None
                )

        _write_shortfall_notes(workbook, analysis, committed)
        output = BytesIO()
        workbook.save(output)
        return output.getvalue()
    finally:
        workbook.close()
