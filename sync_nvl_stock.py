"""Legacy compatibility facade for :mod:`nvl.stock`.

Use `python -m nvl.stock` for the canonical CLI.
"""

from nvl.stock import (
    NVLConfig,
    NVLReconcileResult,
    check_target_sheet_safety,
    generate_nvl_error_report,
    generate_nvl_report,
    identify_sharepoint_metadata_exemption,
    load_nvl_config,
    main,
    normalize_nvl_code,
    parse_nvl_quantity,
    patch_nvl_destination_workbook,
    read_nvl_source_stock,
    reconcile_nvl_target,
    run_nvl_sync,
    verify_nvl_patched_workbook,
)

__all__ = [
    "NVLConfig",
    "NVLReconcileResult",
    "load_nvl_config",
    "normalize_nvl_code",
    "parse_nvl_quantity",
    "check_target_sheet_safety",
    "read_nvl_source_stock",
    "reconcile_nvl_target",
    "patch_nvl_destination_workbook",
    "identify_sharepoint_metadata_exemption",
    "verify_nvl_patched_workbook",
    "generate_nvl_report",
    "generate_nvl_error_report",
    "run_nvl_sync",
]


if __name__ == "__main__":
    main()
