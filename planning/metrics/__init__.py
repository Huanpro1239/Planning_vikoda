"""Canonical Planning metrics API.

Production wiring is static: No kho debt, Danh_muc leadtime and all-month
calculation are explicit dependencies. Importing this package never rebinds
functions or mutable policy state.
"""

from .calculation import (
    bootstrap_opening_debt,
    calculate_metrics,
    calculate_row,
    calculate_row_default,
    clamp_nonnegative_production,
    demand_days,
    excel_roundup_integer,
    planning_stock,
    resolve_plan_year,
    urgent_supply_need,
)
from .constants import (
    DEBT_SHEET,
    FC_SELECTOR_CELL,
    FC_SHEET,
    MASTER_SHEET,
    OUTPUT_COLUMNS,
    PLANNING_SHEET,
)
from .readers import (
    hash_debt_sheet,
    read_actual_inputs,
    read_conversion_factors_and_leadtime,
    read_debt_from_no_kho,
    read_leadtime_from_master,
    read_planning_rows,
    read_planning_rows_robust,
    read_system_receipts,
)
from .service import main, main_with_retry
from .state import (
    RUNTIME_STATE_FILE,
    RUNTIME_STATE_VERSION,
    is_month_end,
    load_runtime_state,
    month_key,
    next_month,
    save_runtime_state,
)
from .workbook import count_changes, patch_workbook

# Temporary compatibility aliases for existing internal call sites/tests.
_bootstrap_opening_debt = bootstrap_opening_debt
_is_month_end = is_month_end
_month_key = month_key
_next_month = next_month
_resolve_plan_year = resolve_plan_year

__all__ = [
    "DEBT_SHEET",
    "FC_SELECTOR_CELL",
    "FC_SHEET",
    "MASTER_SHEET",
    "OUTPUT_COLUMNS",
    "PLANNING_SHEET",
    "RUNTIME_STATE_FILE",
    "RUNTIME_STATE_VERSION",
    "bootstrap_opening_debt",
    "calculate_metrics",
    "calculate_row",
    "calculate_row_default",
    "clamp_nonnegative_production",
    "count_changes",
    "demand_days",
    "excel_roundup_integer",
    "hash_debt_sheet",
    "load_runtime_state",
    "main",
    "main_with_retry",
    "patch_workbook",
    "planning_stock",
    "read_actual_inputs",
    "read_conversion_factors_and_leadtime",
    "read_debt_from_no_kho",
    "read_leadtime_from_master",
    "read_planning_rows",
    "read_planning_rows_robust",
    "read_system_receipts",
    "resolve_plan_year",
    "save_runtime_state",
    "urgent_supply_need",
]
