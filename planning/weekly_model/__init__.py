"""Canonical weekly planning adapter facade.

The pure calculation engine remains in planning.weekly_engine.
"""

from .inputs import (
    MASTER_SHEET,
    PLANNING_SHEET,
    START_COLUMN,
    compute_planning_inputs_hash,
)
from .policy import (
    DEBT_HEADER_NAMES,
    PROFILE_HEADER_NAMES,
    find_header_col,
    normalize_debt_mode,
    normalize_profile,
)
from .report import (
    BALANCE_EPS,
    build_weekly_schedule_report,
    mass_balance,
    validate_shared_machine,
)
from .schedule import (
    ENGINE_VERSION,
    WeeklyAnalysis,
    analyze_weekly_workbook,
)
from .service import prepare_weekly_schedule_update
from .verification import verify_weekly_workbook
from .workbook import patch_weekly_workbook

# Compatibility aliases for historical internal tests/callers.
_mass_balance = mass_balance
_validate_shared_machine = validate_shared_machine

__all__ = [
    "BALANCE_EPS",
    "DEBT_HEADER_NAMES",
    "ENGINE_VERSION",
    "MASTER_SHEET",
    "PLANNING_SHEET",
    "PROFILE_HEADER_NAMES",
    "START_COLUMN",
    "WeeklyAnalysis",
    "analyze_weekly_workbook",
    "build_weekly_schedule_report",
    "compute_planning_inputs_hash",
    "find_header_col",
    "normalize_debt_mode",
    "normalize_profile",
    "patch_weekly_workbook",
    "prepare_weekly_schedule_update",
    "verify_weekly_workbook",
]
