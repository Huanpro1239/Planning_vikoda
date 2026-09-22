"""Weekly schedule report, mass-balance and resource summaries."""

from __future__ import annotations

import hashlib
from collections import defaultdict
from datetime import date

from planning.weekly_engine import PlannerPolicy

from .inputs import EPS
from .schedule import WeeklyAnalysis


BALANCE_EPS = 1e-5


def mass_balance(
    analysis: WeeklyAnalysis,
):
    """Mass balance with explicit service and safety-stock layers."""
    scheduled: dict[str, float] = defaultdict(float)
    for item in analysis.daily_plan:
        scheduled[str(item.ma_sp)] += float(item.qty)

    mass = {}
    service_carryovers = []
    buffer_carryovers = []

    for calc in analysis.calculated:
        code = str(calc.input.ma_sp)
        desired = calc.schedulable_qty
        service_target = calc.service_qty
        done = scheduled.get(code, 0.0)

        if done > desired + BALANCE_EPS:
            raise RuntimeError(
                f"Mã {code} scheduled={done} vượt Q={desired}."
            )

        service_done = min(
            done,
            service_target,
        )
        service_carry = max(
            0.0,
            service_target - service_done,
        )
        buffer_target = calc.buffer_qty
        buffer_done = max(
            0.0,
            done - service_target,
        )
        buffer_done = min(
            buffer_done,
            buffer_target,
        )
        buffer_carry = max(
            0.0,
            buffer_target - buffer_done,
        )
        total_carry = max(
            0.0,
            desired - done,
        )

        if service_carry <= BALANCE_EPS:
            service_carry = 0.0
        if buffer_carry <= BALANCE_EPS:
            buffer_carry = 0.0
        if total_carry <= BALANCE_EPS:
            total_carry = 0.0

        if service_carry:
            service_carryovers.append(code)
        if buffer_carry:
            buffer_carryovers.append(code)

        mass[code] = {
            "uom": calc.input.don_vi_tinh,
            "planned_qty": desired,
            "scheduled_qty": done,
            "carryover_qty": total_carry,
            "balanced_qty": done + total_carry,
            "service_target_qty": service_target,
            "service_scheduled_qty": service_done,
            "service_carryover_qty": service_carry,
            "buffer_target_qty": buffer_target,
            "buffer_scheduled_qty": buffer_done,
            "buffer_carryover_qty": buffer_carry,
        }

    return (
        mass,
        sorted(service_carryovers),
        sorted(buffer_carryovers),
    )


def validate_shared_machine(
    analysis: WeeklyAnalysis,
    policy: PlannerPolicy | None = None,
):
    """Validate physical capacity of the shared KHS + PET 9000 machine."""
    active_policy = policy or PlannerPolicy()
    selected = [
        calc
        for calc in analysis.calculated
        if calc.input.chuyen
        in active_policy.serialized_lines
        and calc.q_rounded > EPS
    ]

    if not selected:
        return {
            "ok": True,
            "state": "not_used",
            "resource": "KHS/PET 9000 shared machine",
            "capacity_shifts_per_day": 0.0,
            "peak_usage_shifts": 0.0,
            "over_capacity_dates": [],
        }

    capacity = selected[0].input.shifts_per_day
    if capacity <= EPS:
        return {
            "ok": False,
            "state": "invalid_capacity",
            "resource": "KHS/PET 9000 shared machine",
            "capacity_shifts_per_day": capacity,
            "peak_usage_shifts": 0.0,
            "over_capacity_dates": [],
        }

    if any(
        abs(
            calc.input.shifts_per_day
            - capacity
        )
        > EPS
        for calc in selected[1:]
    ):
        return {
            "ok": False,
            "state": "inconsistent_shift_calendar",
            "resource": "KHS/PET 9000 shared machine",
            "capacity_shifts_per_day": capacity,
            "peak_usage_shifts": 0.0,
            "over_capacity_dates": [],
        }

    by_code = {
        calc.input.ma_sp: calc
        for calc in selected
    }
    usage: dict[date, float] = defaultdict(float)

    for item in analysis.daily_plan:
        calc = by_code.get(item.ma_sp)
        if calc is not None:
            usage[item.date] += (
                float(item.qty)
                / calc.input.sl_ca
            )

    peak = max(
        usage.values(),
        default=0.0,
    )
    over = sorted(
        current.isoformat()
        for current, used in usage.items()
        if used > capacity + EPS
    )

    return {
        "ok": not over,
        "state": (
            "shared_machine_passed"
            if not over
            else "shared_machine_over_capacity"
        ),
        "resource": "KHS/PET 9000 shared machine",
        "capacity_shifts_per_day": capacity,
        "peak_usage_shifts": peak,
        "over_capacity_dates": over,
    }


def build_weekly_schedule_report(
    workbook_bytes: bytes,
    analysis: WeeklyAnalysis,
    *,
    input_revision=None,
):
    (
        mass,
        service_carryovers,
        buffer_carryovers,
    ) = mass_balance(analysis)

    lines = sorted({
        calc.input.chuyen
        for calc in analysis.calculated
    })
    shared_machine = validate_shared_machine(
        analysis
    )
    shared_name = shared_machine["resource"]

    is_capacity_balanced = any(
        getattr(
            item,
            "phase",
            "",
        )
        == "capacity_balanced"
        for item in analysis.daily_plan
    )

    if analysis.policy_warnings:
        service_state = "policy_metadata_missing"
    elif service_carryovers:
        service_state = "stockout_risk"
    else:
        service_state = "rules_complete"

    if not buffer_carryovers:
        safety_state = "complete"
    elif is_capacity_balanced:
        safety_state = "partially_achieved"
    elif service_carryovers:
        safety_state = (
            "not_achieved_with_service_shortfall"
        )
    else:
        safety_state = "partially_achieved"

    monthly_state = (
        "complete"
        if (
            not service_carryovers
            and not buffer_carryovers
        )
        else (
            "capacity_constrained_balanced"
            if is_capacity_balanced
            else (
                "service_complete_buffer_shortfall"
                if not service_carryovers
                else "service_carryover"
            )
        )
    )

    status = {
        "monthly_quantity": {
            "ok": not service_carryovers,
            "state": monthly_state,
            "carryover_skus": sorted(
                set(service_carryovers)
                | set(buffer_carryovers)
            ),
            "service_carryover_skus": service_carryovers,
            "buffer_carryover_skus": buffer_carryovers,
        },
        "resource_validation": {
            "ok": shared_machine["ok"],
            "state": shared_machine["state"],
            "validated_resources": [shared_name],
            "failed_resources": (
                []
                if shared_machine["ok"]
                else [shared_name]
            ),
            "shared_machine": shared_machine,
        },
        "service": {
            "ok": (
                not analysis.policy_warnings
                and not service_carryovers
            ),
            "state": service_state,
            "stockout_skus": service_carryovers,
            "capacity_balanced_skus": (
                service_carryovers
                if is_capacity_balanced
                else []
            ),
        },
        "safety_stock": {
            "ok": not buffer_carryovers,
            "state": safety_state,
            "shortfall_skus": buffer_carryovers,
        },
    }

    publish_ready = (
        status["monthly_quantity"]["ok"]
        and status["resource_validation"]["ok"]
        and status["service"]["ok"]
    )

    resources = {
        line: {
            "meta": {
                "mode": "ke_hoach_sx_tuan_reference_rules"
            }
        }
        for line in lines
    }
    resources[shared_name] = {
        "meta": shared_machine
    }

    return {
        "schema_version": 5,
        "algorithm": "ke_hoach_sx_tuan_v2_service_first",
        "plan_month": (
            f"{analysis.period_year:04d}-"
            f"{analysis.period_month:02d}"
        ),
        "input_revision": dict(
            input_revision or {}
        ),
        "input_sha256": hashlib.sha256(
            workbook_bytes
        ).hexdigest(),
        "mass_balance": mass,
        "resources": resources,
        "inventory": {},
        "policy_warnings": list(
            analysis.policy_warnings
        ),
        "status": status,
        "publish_status": (
            "ready_for_publish"
            if publish_ready
            else "review_required"
        ),
    }
