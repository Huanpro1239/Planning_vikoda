"""Canonical import surface for NVL stock synchronization.

The implementation remains in the legacy root module during migration. This
module deliberately re-exports only the supported business API so new callers
do not depend on accidental globals from the legacy implementation.
"""

from sync_nvl_stock import (
    NVLConfig,
    NVLReconcileResult,
    load_nvl_config,
    normalize_nvl_code,
    parse_nvl_quantity,
    patch_nvl_destination_workbook,
    run_nvl_sync,
    verify_nvl_patched_workbook,
)

__all__ = [
    "NVLConfig",
    "NVLReconcileResult",
    "load_nvl_config",
    "normalize_nvl_code",
    "parse_nvl_quantity",
    "patch_nvl_destination_workbook",
    "run_nvl_sync",
    "verify_nvl_patched_workbook",
]
