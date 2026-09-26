"""Build a concise operational summary from NVL run artifacts."""

from __future__ import annotations

from datetime import datetime, timezone
import json
from pathlib import Path
from typing import Any


SUMMARY_SCHEMA = "nvl_operational_summary_v1"


def _load_optional(path: Path) -> dict[str, Any] | None:
    if not path.is_file():
        return None
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"_invalid_json": True, "_path": str(path)}
    return value if isinstance(value, dict) else {"_invalid_json": True}


def build_operational_summary(
    root: str | Path = ".",
    *,
    job_status: str = "unknown",
    event_name: str = "",
    run_id: str = "",
    commit_sha: str = "",
) -> dict[str, Any]:
    root = Path(root)
    readiness = _load_optional(root / "nvl_production_readiness_report.json")
    stock = _load_optional(root / "nvl_stock_report.json")
    open_po = _load_optional(root / "nvl_open_po_report.json")
    manifest = _load_optional(root / "nvl_release_manifest.json")
    ledger = _load_optional(root / "nvl_release_index.json")
    staging = _load_optional(root / "staging_copy_info.json")
    recovery = _load_optional(root / "nvl_recovery_report.json")

    readiness_status = (
        readiness.get("status") if readiness else "missing"
    )
    stock_status = stock.get("status") if stock else "missing"
    stock_mode = stock.get("mode") if stock else None
    stock_metrics = (stock.get("metrics") or {}) if stock else {}
    open_target = (open_po.get("target") or {}) if open_po else {}
    published_workbook = (
        manifest.get("published_workbook") or {}
        if manifest
        else {}
    )
    upstream_planning = (
        manifest.get("upstream_planning") or {}
        if manifest
        else {}
    )
    ledger_summary = (
        ledger.get("summary") or {}
        if ledger
        else {}
    )

    failed_phase = None
    failure_detail = None
    if readiness and readiness.get("status") == "failed":
        failed_phase = readiness.get("failed_phase") or "readiness"
    elif stock and stock.get("status") == "failed":
        failed_phase = stock.get("phase") or "stock"
        failure_detail = stock.get("message")
    elif open_po and (
        open_po.get("status") == "failed"
        or open_po.get("error")
    ):
        failed_phase = open_po.get("phase") or "open_po"
        failure_detail = open_po.get("message") or open_po.get("error")
    elif ledger and ledger.get("status") == "failed":
        failed_phase = "release_ledger"
    elif recovery and recovery.get("status") == "failed":
        failed_phase = "recovery"

    normalized_job = str(job_status or "unknown").strip().lower()
    outcome = "PASS" if normalized_job == "success" else (
        "FAIL" if normalized_job in {"failure", "cancelled", "timed_out"} else "UNKNOWN"
    )

    return {
        "schema": SUMMARY_SCHEMA,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "outcome": outcome,
        "job_status": normalized_job,
        "event_name": event_name or None,
        "run_id": run_id or None,
        "commit_sha": commit_sha or None,
        "failed_phase": failed_phase,
        "failure_detail": failure_detail,
        "readiness": {
            "status": readiness_status,
            "gate_version": readiness.get("gate_version") if readiness else None,
        },
        "stock": {
            "mode": stock_mode,
            "status": stock_status,
            "changed_D": stock_metrics.get("changed_count"),
            "target_revision": (
                (stock.get("target") or {}).get("revision")
                if stock else None
            ),
        },
        "open_po": {
            "published": open_po.get("published") if open_po else None,
            "changed_E": open_target.get("changed_cells"),
            "target_revision_after": (
                (open_po.get("publish_evidence") or {}).get(
                    "target_revision_after"
                )
                if open_po else None
            ),
        },
        "release": {
            "release_id": manifest.get("release_id") if manifest else None,
            "release_version": (
                manifest.get("release_version") if manifest else None
            ),
            "workbook_sha256": published_workbook.get("sha256"),
            "upstream_planning_run_id": upstream_planning.get("run_id"),
            "upstream_planning_head_sha": upstream_planning.get("head_sha"),
            "upstream_planning_event": upstream_planning.get("event"),
        },
        "ledger": {
            "status": ledger.get("status") if ledger else "missing",
            "release_count": ledger_summary.get("release_count"),
            "failed_count": ledger_summary.get("failed_count"),
            "warning_count": ledger_summary.get("warning_count"),
        },
        "staging": {
            "status": staging.get("status") if staging else None,
            "name": staging.get("name") if staging else None,
        },
        "recovery": {
            "status": recovery.get("status") if recovery else None,
            "release_id": recovery.get("release_id") if recovery else None,
            "changed_cells": (
                (recovery.get("proposal") or {}).get("changed_cells")
                if recovery else None
            ),
        },
    }


def render_markdown(summary: dict[str, Any]) -> str:
    outcome = summary.get("outcome")
    readiness = summary.get("readiness") or {}
    stock = summary.get("stock") or {}
    open_po = summary.get("open_po") or {}
    release = summary.get("release") or {}
    ledger = summary.get("ledger") or {}

    lines = [
        "## NVL Operational Summary",
        "",
        f"**Outcome:** {outcome}",
        "",
        "| Signal | Value |",
        "|---|---|",
        f"| Readiness | {readiness.get('status')} |",
        f"| Stock D | status={stock.get('status')}; changed={stock.get('changed_D')} |",
        f"| Open PO E | published={open_po.get('published')}; changed={open_po.get('changed_E')} |",
        f"| Release ID | {release.get('release_id') or '-'} |",
        f"| Upstream Planning run | {release.get('upstream_planning_run_id') or '-'} |",
        f"| Upstream Planning event | {release.get('upstream_planning_event') or '-'} |",
        f"| Workbook SHA-256 | {release.get('workbook_sha256') or '-'} |",
        f"| Ledger | status={ledger.get('status')}; releases={ledger.get('release_count')} |",
    ]
    if summary.get("failed_phase"):
        lines.extend(
            [
                "",
                f"**Failed phase:** {summary.get('failed_phase')}",
            ]
        )
        if summary.get("failure_detail"):
            lines.append(
                f"**Failure:** {summary.get('failure_detail')}"
            )
    return "\n".join(lines) + "\n"


def write_operational_summary(
    summary: dict[str, Any],
    *,
    json_path: str | Path = "nvl_operational_summary.json",
    markdown_path: str | Path | None = None,
) -> None:
    Path(json_path).write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True)
        + "\n",
        encoding="utf-8",
    )
    if markdown_path:
        path = Path(markdown_path)
        with path.open("a", encoding="utf-8") as handle:
            handle.write(render_markdown(summary))
