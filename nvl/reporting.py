"""NVL audit and failure reports."""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from nvl.models import NVLConfig, NVLReconcileResult

def generate_nvl_report(
    reconcile_result: NVLReconcileResult,
    config: NVLConfig,
    *,
    mode: str,
    source_revision: str | None = None,
    target_revision: str | None = None,
    reporting_period: str | None = None,
    exempted_parts: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Tạo báo cáo JSON đối soát chi tiết đồng bộ tồn NVL."""
    return {
        "schema_version": 1,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "mode": mode,
        "status": reconcile_result.status,
        "message": reconcile_result.message,
        "source": {
            "name": config.source_name,
            "sharepoint_path": config.source_path,
            "sourcedoc": config.source_sourcedoc,
            "sheet_name": config.source_sheet,
            "code_column": config.source_code_col_letter,
            "value_column": config.source_value_col_letter,
            "revision": source_revision,
            "reporting_period": reporting_period,
        },
        "target": {
            "name": config.target_name,
            "sharepoint_path": config.target_path,
            "sourcedoc": config.target_sourcedoc,
            "sheet_name": config.target_sheet,
            "code_column": config.target_code_col_letter,
            "value_column": config.target_value_col_letter,
            "revision": target_revision,
        },
        "metrics": {
            "matched_count": len(reconcile_result.changes) + len(reconcile_result.unchanged),
            "changed_count": len(reconcile_result.changes),
            "unchanged_count": len(reconcile_result.unchanged),
            "missing_in_source_count": len(reconcile_result.missing_in_source),
            "source_only_count": len(reconcile_result.source_only),
        },
        "changes": reconcile_result.changes,
        "warnings": [
            {
                "type": "missing_in_source",
                "row": item["row"],
                "code": item["code"],
                "message": f"Mã vật tư '{item['code']}' ở đích không có trong file nguồn; giữ nguyên số tồn cũ ({item['current_value']}).",
            }
            for item in reconcile_result.missing_in_source
        ],
        "source_only_codes": reconcile_result.source_only,
        "exempted_server_metadata_parts": exempted_parts or [],
    }


def generate_nvl_error_report(
    config: NVLConfig | None = None,
    *,
    mode: str = "unknown",
    phase: str,
    attempt: int,
    error: Exception,
    source_revision: str | None = None,
    target_revision: str | None = None,
    extra: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Tạo báo cáo lỗi JSON khi đồng bộ hoặc publish thất bại."""
    src_name = getattr(config, "source_name", "XNT_ketoan_Vikoda.xlsm") if config else "XNT_ketoan_Vikoda.xlsm"
    src_path = getattr(config, "source_path", "") if config else ""
    src_doc = getattr(config, "source_sourcedoc", "") if config else ""

    tgt_name = getattr(config, "target_name", "Kế hoạch mua hàng.xlsx") if config else "Kế hoạch mua hàng.xlsx"
    tgt_path = getattr(config, "target_path", "") if config else ""
    tgt_doc = getattr(config, "target_sourcedoc", "") if config else ""

    rep: dict[str, Any] = {
        "schema_version": 1,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "mode": mode,
        "status": "failed",
        "phase": phase,
        "attempt": attempt,
        "error_type": type(error).__name__,
        "message": str(error),
        "error_message": str(error),
        "source": {
            "name": src_name,
            "sharepoint_path": src_path,
            "sourcedoc": src_doc,
            "revision": source_revision,
        },
        "target": {
            "name": tgt_name,
            "sharepoint_path": tgt_path,
            "sourcedoc": tgt_doc,
            "revision": target_revision,
        },
    }
    if extra:
        rep.update(extra)
    return rep


__all__ = ["generate_nvl_report", "generate_nvl_error_report"]
