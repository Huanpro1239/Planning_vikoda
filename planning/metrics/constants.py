"""Shared constants for Planning metrics."""

FC_SELECTOR_CELL = "R1"
FC_SHEET = "FC"
PLANNING_SHEET = "Ke_hoach_SX"
MASTER_SHEET = "Danh_muc"
MASTER_LEADTIME_COL = 10

DEBT_SHEET = "No kho"
DEBT_CODE_COL = 1
DEBT_VALUE_COL = 4

ACTUAL_CODE_COL = 3
ACTUAL_RECEIPT_COL = 7
ACTUAL_CONSIGNMENT_COL = 17

SYSTEM_CODE_COL = 2
SYSTEM_RECEIPT_COL = 9

COL_BATCH = 4
COL_PER_SHIFT = 5
COL_CLASSIFICATION = 8
COL_SHIFTS_PER_DAY = 9
COL_ACTUAL_STOCK = 10
COL_BOOK_STOCK = 11
COL_FC = 12
COL_CURRENT_DEBT = 14

OUTPUT_COLUMNS = {
    "M": "expected_end_stock",
    "N": "warehouse_debt",
    "O": "required_production",
    "P": "rounded_production",
    "Q": "production_days",
    "R": "production_start",
}
