"""Independent verification for the weekly planning adapter."""

from __future__ import annotations

import math
from typing import Any

from .report import (
    BALANCE_EPS,
    mass_balance,
    validate_shared_machine,
)
from .schedule import (
    ENGINE_VERSION,
    analyze_weekly_workbook,
)


def verify_weekly_workbook(
    workbook_bytes: bytes,
    *,
    schedule_report: dict[str, Any],
    plan_year: int,
    plan_month: int,
):
    revision = schedule_report.get(
        "input_revision"
    )
    if (
        not isinstance(revision, dict)
        or not revision
    ):
        raise RuntimeError(
            "Weekly schedule report thiếu "
            "input_revision/provenance."
        )

    if schedule_report.get("plan_month") != (
        f"{plan_year:04d}-{plan_month:02d}"
    ):
        raise RuntimeError(
            "Weekly schedule report sai kỳ kế hoạch."
        )

    analysis = analyze_weekly_workbook(
        workbook_bytes,
        plan_year=plan_year,
        plan_month=plan_month,
    )
    if analysis.changed_cells:
        raise RuntimeError(
            "Workbook chưa khớp weekly engine: "
            f"{analysis.changed_cells} ô sai."
        )

    (
        expected_mass,
        service_carryovers,
        buffer_carryovers,
    ) = mass_balance(analysis)
    actual_mass = (
        schedule_report.get("mass_balance")
        or {}
    )

    for code, expected in expected_mass.items():
        actual = actual_mass.get(code) or {}
        for key in (
            "planned_qty",
            "scheduled_qty",
            "carryover_qty",
            "service_target_qty",
            "service_scheduled_qty",
            "service_carryover_qty",
            "buffer_target_qty",
            "buffer_scheduled_qty",
            "buffer_carryover_qty",
        ):
            if not math.isclose(
                float(
                    actual.get(key, 0)
                    or 0
                ),
                float(expected[key]),
                rel_tol=1e-9,
                abs_tol=BALANCE_EPS,
            ):
                raise RuntimeError(
                    f"Mass balance {code}.{key} "
                    "không khớp workbook."
                )

    expected_warnings = sorted(
        (
            warning["code"],
            warning["type"],
        )
        for warning
        in analysis.policy_warnings
    )
    report_warnings = sorted(
        (
            str(warning.get("code")),
            str(warning.get("type")),
        )
        for warning
        in schedule_report.get(
            "policy_warnings",
            [],
        )
        if isinstance(warning, dict)
    )
    if expected_warnings != report_warnings:
        raise RuntimeError(
            "policy_warnings không khớp workbook."
        )

    expected_resource = validate_shared_machine(
        analysis
    )
    report_resource = (
        (
            schedule_report.get("status")
            or {}
        ).get("resource_validation")
        or {}
    )
    if (
        bool(report_resource.get("ok"))
        != bool(expected_resource["ok"])
        or report_resource.get("state")
        != expected_resource["state"]
    ):
        raise RuntimeError(
            "resource_validation không khớp "
            "shared-machine schedule."
        )

    if (
        service_carryovers
        and schedule_report.get(
            "publish_status"
        )
        in (
            "ready_for_publish",
            "feasible",
        )
    ):
        raise RuntimeError(
            "Kế hoạch còn thiếu service "
            f"({service_carryovers}) nhưng "
            "publish_status lại là "
            f"{schedule_report.get('publish_status')}."
        )

    return {
        "validated": True,
        "algorithm": ENGINE_VERSION,
        "checked_skus": len(
            analysis.calculated
        ),
        "service_carryover_skus": service_carryovers,
        "buffer_carryover_skus": buffer_carryovers,
        "policy_warning_count": len(
            analysis.policy_warnings
        ),
        "publish_status": schedule_report.get(
            "publish_status"
        ),
    }
