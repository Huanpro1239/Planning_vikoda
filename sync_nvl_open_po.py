"""Legacy compatibility CLI for NVL open-PO synchronization.

Business logic lives in :mod:`nvl.open_po`. Keep this module thin because
existing workflows and external callers may still execute/import it directly.
"""

from nvl.open_po import (
    DEFAULT_PROPOSAL,
    DEFAULT_REPORT,
    DEFAULT_SOURCE_CONFIG,
    DEFAULT_TARGET_CONFIG,
    EPS,
    OpenPOConfig,
    build_report,
    load_open_po_config,
    main,
    patch_target_workbook,
    read_open_po,
    reconcile_target,
    resolve_source_item,
    run_online,
    verify_target,
)

__all__ = [
    "OpenPOConfig",
    "load_open_po_config",
    "patch_target_workbook",
    "read_open_po",
    "reconcile_target",
    "run_online",
    "verify_target",
]


if __name__ == "__main__":
    main()
