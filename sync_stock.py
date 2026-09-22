"""Legacy CLI/facade for finished-goods stock synchronization.

Canonical business logic lives in the stock package. This root module is kept
for backward-compatible imports and direct CLI execution only.
"""

from excel.workbook_xml import (
    column_number,
    find_sheet_xml_path,
    load_shared_strings,
    read_cell_text,
    set_numeric_cell,
)
from sharepoint.client import (
    GRAPH,
    HOSTNAME,
    SITE_PATH,
    GraphClient,
    GraphRequestError,
    get_access_token,
    is_retryable_graph_error,
)
from stock import (
    DEST_PATH,
    DEST_SHEET,
    MASTER_SHEET,
    SOURCE_ACCOUNTING_VIKODA_PATH,
    SOURCE_ACCOUNTING_VKD_PATH,
    SOURCE_ACTUAL_PATH,
    SOURCE_FACTORY_VIKODA_PATH,
    SOURCE_FACTORY_VKD_PATH,
    SYNC_VERSION,
    clean_number,
    normalize_code,
    patch_destination_workbook,
    read_actual_stock,
    read_conversion_factors,
    read_single_value_source,
    to_number,
)
from stock import state as _state
from stock.service import main


STATE_FILE = _state.STATE_FILE


def load_state():
    return _state.load_state(path=STATE_FILE)


def save_state(
    source_etags,
    conversion_hash,
    fc_hash=None,
    no_kho_hash=None,
    planning_inputs_hash=None,
    engine_version=None,
    **kwargs,
):
    return _state.save_state(
        source_etags,
        conversion_hash,
        fc_hash=fc_hash,
        no_kho_hash=no_kho_hash,
        planning_inputs_hash=planning_inputs_hash,
        engine_version=engine_version,
        path=STATE_FILE,
        **kwargs,
    )


__all__ = [
    "DEST_PATH",
    "DEST_SHEET",
    "GRAPH",
    "HOSTNAME",
    "MASTER_SHEET",
    "SITE_PATH",
    "SOURCE_ACCOUNTING_VIKODA_PATH",
    "SOURCE_ACCOUNTING_VKD_PATH",
    "SOURCE_ACTUAL_PATH",
    "SOURCE_FACTORY_VIKODA_PATH",
    "SOURCE_FACTORY_VKD_PATH",
    "STATE_FILE",
    "SYNC_VERSION",
    "GraphClient",
    "GraphRequestError",
    "clean_number",
    "column_number",
    "find_sheet_xml_path",
    "get_access_token",
    "is_retryable_graph_error",
    "load_shared_strings",
    "load_state",
    "normalize_code",
    "patch_destination_workbook",
    "read_actual_stock",
    "read_cell_text",
    "read_conversion_factors",
    "read_single_value_source",
    "save_state",
    "set_numeric_cell",
    "to_number",
]


if __name__ == "__main__":
    main()
