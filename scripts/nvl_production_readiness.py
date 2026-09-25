"""Single production-readiness gate for NVL releases.

Usage:
    python -X utf8 scripts/nvl_production_readiness.py

A production NVL run is ready only when every phase passes.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import sys
import time
from datetime import datetime, timezone


ROOT = Path(__file__).resolve().parents[1]
REPORT_PATH = ROOT / "nvl_production_readiness_report.json"
GATE_VERSION = "nvl_production_readiness_v1"


def _command(*args: str) -> list[str]:
    return [sys.executable, *args]


PHASES = (
    (
        "architecture",
        _command(
            "-m",
            "unittest",
            "-v",
            "tests.test_package_layout",
            "tests.test_repo_hygiene",
            "tests.test_nvl_workflow_contract",
            "tests.test_oidc_workflow_contract.OIDCWorkflowContractTests.test_nvl_workflow_is_oidc_only",
        ),
    ),
    (
        "business_contracts",
        _command(
            "-m",
            "unittest",
            "-v",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_normalize_nvl_code",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_parse_nvl_quantity",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_read_nvl_source_stock_success",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_read_source_stock_rejects_duplicates",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_missing_in_source_with_text_and_blank_preserved_safely",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_no_matching_codes_raises_error",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_preserve_missing_in_source_false_raises_error",
            "tests.test_sync_nvl_open_po.OpenPOTests.test_read_open_po_filters_and_clamps_per_line",
            "tests.test_sync_nvl_open_po.OpenPOTests.test_reconcile_sets_missing_code_to_zero",
        ),
    ),
    (
        "workbook_roundtrip",
        _command(
            "-m",
            "unittest",
            "-v",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_readers_handle_missing_and_understated_dimensions",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_unmodified_cells_in_ton_nvl_verified_without_corruption",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_reconcile_and_patch_nvl_workbook_full_flow",
            "tests.test_sync_nvl_stock.SyncNVLStockTests.test_idempotence_second_run_has_zero_changes",
            "tests.test_sync_nvl_open_po.OpenPOTests.test_patch_only_e_and_verify",
        ),
    ),
    (
        "deterministic_proposal",
        _command(
            "-m",
            "unittest",
            "-v",
            "tests.nvl_readiness.test_release_safety.NVLReleaseSafetyTests.test_same_inputs_produce_identical_proposal",
        ),
    ),
    (
        "etag_concurrency",
        _command(
            "-m",
            "unittest",
            "-v",
            "tests.test_nvl_publish_boundary.NVLPublishBoundaryTests.test_publish_with_changes_calls_upload_with_expected_etag",
            "tests.test_nvl_publish_boundary.NVLPublishBoundaryTests.test_http_412_refetches_and_retries_upload_with_new_etag",
            "tests.test_nvl_publish_boundary.NVLPublishBoundaryTests.test_source_changed_after_download_triggers_retry_and_uploads_latest",
            "tests.test_nvl_publish_boundary.NVLPublishBoundaryTests.test_timeout_server_committed_concurrent_edit_in_d_stops_without_reupload",
            "tests.test_nvl_publish_boundary.NVLPublishBoundaryTests.test_timeout_server_committed_concurrent_edit_outside_d_stops_without_reupload",
            "tests.test_nvl_publish_boundary.NVLPublishBoundaryTests.test_backup_write_failure_stops_with_zero_upload_attempts",
        ),
    ),
    (
        "publish_dry_run",
        _command(
            "-m",
            "unittest",
            "-v",
            "tests.test_nvl_publish_boundary.NVLPublishBoundaryTests.test_dry_run_does_not_upload",
            "tests.test_nvl_publish_boundary.NVLPublishBoundaryTests.test_publish_without_changes_skips_upload",
            "tests.nvl_readiness.test_release_safety.NVLReleaseSafetyTests.test_dry_run_produces_auditable_artifacts_without_upload",
            "tests.nvl_readiness.test_release_manifest",
            "tests.nvl_readiness.test_release_verification",
            "tests.nvl_readiness.test_release_audit",
        ),
    ),
    (
        "full_regression",
        _command(
            "-m",
            "unittest",
            "discover",
            "-s",
            "tests",
            "-t",
            ".",
            "-p",
            "test_*.py",
            "-v",
        ),
    ),
)


def _write_report(report: dict) -> None:
    REPORT_PATH.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def main() -> int:
    report = {
        "schema_version": 1,
        "gate_version": GATE_VERSION,
        "status": "running",
        "started_at": datetime.now(timezone.utc).isoformat(),
        "git_sha": os.getenv("GITHUB_SHA") or "",
        "python": sys.version,
        "phases": [],
    }
    _write_report(report)

    print(f"[NVL-READINESS] gate={GATE_VERSION}")
    for name, command in PHASES:
        started = time.monotonic()
        print(f"[NVL-READINESS][{name}] START")
        result = subprocess.run(command, cwd=ROOT, check=False)
        duration = round(time.monotonic() - started, 3)
        phase = {
            "name": name,
            "status": "passed" if result.returncode == 0 else "failed",
            "returncode": result.returncode,
            "duration_seconds": duration,
            "command": command,
        }
        report["phases"].append(phase)
        _write_report(report)

        if result.returncode != 0:
            report["status"] = "failed"
            report["failed_phase"] = name
            report["finished_at"] = datetime.now(timezone.utc).isoformat()
            _write_report(report)
            print(
                f"[NVL-READINESS][{name}] FAIL rc={result.returncode}; "
                f"report={REPORT_PATH.name}"
            )
            return result.returncode

        print(f"[NVL-READINESS][{name}] PASS ({duration:.3f}s)")

    report["status"] = "passed"
    report["finished_at"] = datetime.now(timezone.utc).isoformat()
    _write_report(report)
    print(f"[NVL-READINESS] PASS report={REPORT_PATH.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
