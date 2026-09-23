"""Immutable release manifest for successful Planning publishes."""

from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from planning_schedule_report import json_safe

from .proposal import proposal_output_sha256


RELEASE_MANIFEST_FILE = Path("planning_release_manifest.json")
READINESS_REPORT_FILE = Path("production_readiness_report.json")
RELEASE_SCHEMA_VERSION = 1
RELEASE_SCHEMA = "planning_release_manifest_v1"


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _read_readiness_report(
    path: Path = READINESS_REPORT_FILE,
    *,
    required: bool = False,
) -> dict[str, Any]:
    if not path.exists():
        if required:
            raise RuntimeError(
                "Production publish thiếu production_readiness_report.json."
            )
        return {
            "status": "not_available",
            "gate_version": None,
            "git_sha": None,
        }

    try:
        report = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise RuntimeError(
            "Không đọc được production readiness report."
        ) from exc

    if required and report.get("status") != "passed":
        raise RuntimeError(
            "Production readiness gate chưa PASS; không tạo release."
        )

    return report


def build_release_manifest(
    report: dict[str, Any],
    decision: dict[str, Any],
    final_bytes: bytes,
    *,
    published_at: str | None = None,
    environ: dict[str, str] | None = None,
    readiness_path: Path = READINESS_REPORT_FILE,
) -> dict[str, Any]:
    env = dict(os.environ if environ is None else environ)
    strict = str(env.get("PLANNING_REQUIRE_READINESS", "")).strip() in {
        "1",
        "true",
        "TRUE",
        "yes",
        "YES",
    }
    readiness = _read_readiness_report(
        readiness_path,
        required=strict,
    )

    commit_sha = (
        str(env.get("GITHUB_SHA") or "").strip()
        or str(readiness.get("git_sha") or "").strip()
    )
    if strict and not commit_sha:
        raise RuntimeError(
            "Production publish thiếu commit SHA để truy vết release."
        )

    proposal_id = str(report.get("proposal_id") or "").strip()
    if not proposal_id:
        raise RuntimeError("Release manifest thiếu proposal_id.")

    publish_state = str(decision.get("state") or "").strip()
    if publish_state not in {"published", "published_no_change"}:
        raise RuntimeError(
            "Chỉ publish thành công mới được tạo release manifest."
        )

    plan_month = str(report.get("plan_month") or "unknown")
    run_id = str(env.get("GITHUB_RUN_ID") or "").strip()
    run_attempt = str(env.get("GITHUB_RUN_ATTEMPT") or "1").strip()
    timestamp = published_at or datetime.now(timezone.utc).isoformat()

    instance = (
        f"gh-{run_id}-{run_attempt}"
        if run_id
        else timestamp.replace(":", "").replace("+", "_")
    )
    release_id = (
        f"{plan_month}-{proposal_id[:12]}-{instance}"
    )
    release_version = (
        f"{RELEASE_SCHEMA}.proposal-{proposal_id[:12]}"
        + (f".run-{run_id}.{run_attempt}" if run_id else "")
    )

    pipeline = report.get("pipeline") or {}
    manifest = {
        "schema": RELEASE_SCHEMA,
        "schema_version": RELEASE_SCHEMA_VERSION,
        "release_id": release_id,
        "release_version": release_version,
        "published_at": timestamp,
        "plan_month": report.get("plan_month"),
        "algorithm": report.get("algorithm"),
        "engine_version": (
            pipeline.get("engine_version")
            or pipeline.get("engine")
        ),
        "commit": {
            "sha": commit_sha or None,
            "ref": env.get("GITHUB_REF") or None,
            "repository": env.get("GITHUB_REPOSITORY") or None,
            "run_id": run_id or None,
            "run_attempt": run_attempt if run_id else None,
            "workflow": env.get("GITHUB_WORKFLOW") or None,
            "actor": env.get("GITHUB_ACTOR") or None,
        },
        "readiness": {
            "status": readiness.get("status"),
            "gate_version": readiness.get("gate_version"),
            "git_sha": readiness.get("git_sha"),
            "report_sha256": (
                _sha256_bytes(readiness_path.read_bytes())
                if readiness_path.exists()
                else None
            ),
        },
        "proposal": {
            "proposal_id": proposal_id,
            "output_sha256": _sha256_bytes(final_bytes),
            "stable_output_sha256": proposal_output_sha256(final_bytes),
        },
        "input_revision": report.get("input_revision") or {},
        "publish_decision": dict(decision),
        "pipeline_fingerprints": {
            key: pipeline.get(key)
            for key in (
                "conversion_hash",
                "fc_hash",
                "no_kho_hash",
                "planning_inputs_hash",
                "engine_version",
            )
            if pipeline.get(key) is not None
        },
    }
    return json_safe(manifest)


def save_release_manifest(
    report: dict[str, Any],
    decision: dict[str, Any],
    final_bytes: bytes,
    *,
    path: Path = RELEASE_MANIFEST_FILE,
) -> dict[str, Any]:
    manifest = build_release_manifest(
        report,
        decision,
        final_bytes,
    )
    path.write_text(
        json.dumps(
            manifest,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
    )
    return manifest
