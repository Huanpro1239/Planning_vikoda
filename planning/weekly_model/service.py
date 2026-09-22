"""Thin weekly-model orchestration."""

from .report import build_weekly_schedule_report
from .schedule import analyze_weekly_workbook
from .workbook import patch_weekly_workbook


def prepare_weekly_schedule_update(
    workbook_bytes: bytes,
    *,
    plan_year: int,
    plan_month: int,
    input_revision=None,
):
    analysis = analyze_weekly_workbook(
        workbook_bytes,
        plan_year=plan_year,
        plan_month=plan_month,
    )
    updated = patch_weekly_workbook(
        workbook_bytes,
        analysis,
    )
    report = build_weekly_schedule_report(
        workbook_bytes,
        analysis,
        input_revision=input_revision,
    )
    return updated, report, analysis
