"""Immutable release manifests for successful Planning publishes."""

from __future__ import annotations

import hashlib
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from planning_schedule_report import json_safe

from .constants import (
    PRODUCTION_READINESS_REPORT_FILE,
    RELEASE_MANIFEST_DIR,
    RELEASE_MANIFEST_FILE,
)
from .proposal import proposal_output_sha256


RELEASE_MANIFEST_VERSION = "planning_release_manifest_v1"


def _canonical_bytes(value: Any) -> bytes:
    return json.dumps(
        json_safe(value),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def _sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _safe_token(value: Any, fallback: str = "unknown") -> str:
    token = re.sub(r"[^A-Za-z0-9._-]+", "-", str(value or "").strip())
    token = token.strip("-.")
    return token or fallback


def load_readiness_evidence(
    path: Path = PRODUCTION_READINESS_REPORT_FILE,
) -> dict[str, Any]:
    """Load auditable evidence from the readiness gate without mutating it."""
    path = Path(path)
    if not path.is_file():
        return {
            "status": "unavailable",
            "gate_version": None,
            "git_sha": None,
            "report_sha256": None,
            "phases": [],
        }

    raw = path.read_bytes()
    try:
        payload = json.loads(raw.decode("utf-8"))
    except Exception as exc:
        raise RuntimeError(
            f"Release readiness report không đọc được: {path}"
        ) from exc

    return {
        "status": payload.get("status"),
        "gate_version": payload.get("gate_version"),
        "git_sha": payload.get("git_sha"),
        "report_sha256": _sha256_bytes(raw),
        "phases": [
            {
                "name": phase.get("name"),
                "status": phase.get("status"),
                "returncode": phase.get("returncode"),
            }
            for phase in payload.get("phases", [])
            if isinstance(phase, dict)
        ],
    }


def build_release_manifest(
    report: dict[str, Any],
    decision: dict[str, Any],
    final_bytes: bytes,
    *,
    uploaded: bool,
    target_before: dict[str, Any] | None = None,
    upload_result: dict[str, Any] | None = None,
    environment: dict[str, str] | None = None,
    published_at: datetime | None = None,
    readiness_evidence: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build one traceable release record after authorized publish completion."""
    env = dict(os.environ if environment is None else environment)
    now = published_at or datetime.now(timezone.utc)
    if now.tzinfo is None:
        now = now.replace(tzinfo=timezone.utc)
    now = now.astimezone(timezone.utc)

    proposal_id = str(report.get("proposal_id") or "").strip()
    if not proposal_id:
        raise RuntimeError("Không thể tạo release manifest khi thiếu proposal_id.")

    plan_month = str(report.get("plan_month") or "unknown")
    git_sha = str(env.get("GITHUB_SHA") or "").strip()
    run_id = str(env.get("GITHUB_RUN_ID") or "").strip()
    run_number = str(env.get("GITHUB_RUN_NUMBER") or run_id or "local").strip()
    run_attempt = str(env.get("GITHUB_RUN_ATTEMPT") or "1").strip()

    version = (
        f"{_safe_token(plan_month)}."
        f"r{_safe_token(run_number, 'local')}."
        f"a{_safe_token(run_attempt, '1')}"
    )
    if run_id:
        run_token = f"run{_safe_token(run_id)}-a{_safe_token(run_attempt, '1')}"
    else:
        run_token = (
            f"{_safe_token(git_sha[:12], 'local')}-"
            f"{now.strftime('%Y%m%dT%H%M%SZ')}"
        )
    release_id = (
        f"planning-{_safe_token(plan_month)}-"
        f"{run_token}-{_safe_token(proposal_id[:12])}"
    )

    raw_sha = str(report.get("output_sha256") or _sha256_bytes(final_bytes))
    stable_sha = str(
        report.get("proposal_output_sha256")
        or proposal_output_sha256(final_bytes)
    )
    report_sha = _sha256_bytes(_canonical_bytes(report))
    readiness = (
        dict(readiness_evidence)
        if readiness_evidence is not None
        else load_readiness_evidence()
    )

    target_before = dict(target_before or {})
    upload_result = dict(upload_result or {})
    target_after_etag = (
        upload_result.get("eTag")
        or upload_result.get("etag")
        or target_before.get("eTag")
    )

    manifest = {
        "schema_version": 1,
        "manifest_version": RELEASE_MANIFEST_VERSION,
        "release_id": release_id,
        "release_version": version,
        "published_at": now.isoformat(),
        "plan_month": plan_month,
        "algorithm": report.get("algorithm"),
        "engine_version": (report.get("pipeline") or {}).get("engine_version"),
        "proposal_id": proposal_id,
        "publish": {
            "uploaded": bool(uploaded),
            "decision": dict(decision),
        },
        "source_control": {
            "commit_sha": git_sha or None,
            "repository": env.get("GITHUB_REPOSITORY") or None,
            "ref": env.get("GITHUB_REF") or None,
            "event_name": env.get("GITHUB_EVENT_NAME") or None,
            "workflow": env.get("GITHUB_WORKFLOW") or None,
            "run_id": run_id or None,
            "run_number": env.get("GITHUB_RUN_NUMBER") or None,
            "run_attempt": env.get("GITHUB_RUN_ATTEMPT") or None,
            "actor": env.get("GITHUB_ACTOR") or None,
            "server_url": env.get("GITHUB_SERVER_URL") or None,
        },
        "readiness_gate": readiness,
        "input_revision": dict(report.get("input_revision") or {}),
        "artifact": {
            "name": "planning_proposal.xlsx",
            "size_bytes": len(final_bytes),
            "raw_sha256": raw_sha,
            "stable_sha256": stable_sha,
            "schedule_report_sha256": report_sha,
        },
        "target": {
            "item_id": target_before.get("id"),
            "etag_before": target_before.get("eTag"),
            "etag_after": target_after_etag,
            "last_modified_before": target_before.get("lastModifiedDateTime"),
        },
    }
    manifest["record_sha256"] = _sha256_bytes(_canonical_bytes(manifest))
    return manifest


def validate_release_manifest(manifest: dict[str, Any]) -> None:
    required = (
        "release_id",
        "release_version",
        "proposal_id",
        "input_revision",
        "artifact",
        "readiness_gate",
        "record_sha256",
    )
    missing = [key for key in required if not manifest.get(key)]
    if missing:
        raise RuntimeError(
            "Release manifest thiếu trường bắt buộc: " + ", ".join(missing)
        )

    copy = dict(manifest)
    expected = str(copy.pop("record_sha256"))
    actual = _sha256_bytes(_canonical_bytes(copy))
    if actual != expected:
        raise RuntimeError("Release manifest record_sha256 không hợp lệ.")


def save_release_manifest(
    manifest: dict[str, Any],
    *,
    latest_path: Path = RELEASE_MANIFEST_FILE,
    archive_dir: Path = RELEASE_MANIFEST_DIR,
) -> Path:
    """Persist latest alias plus immutable versioned release record."""
    validate_release_manifest(manifest)

    latest_path = Path(latest_path)
    archive_dir = Path(archive_dir)
    archive_dir.mkdir(parents=True, exist_ok=True)

    release_id = _safe_token(manifest["release_id"])
    immutable_path = archive_dir / f"{release_id}.json"
    payload = (
        json.dumps(
            json_safe(manifest),
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )
        + "\n"
    )

    if immutable_path.exists():
        current = immutable_path.read_text(encoding="utf-8")
        if current != payload:
            raise RuntimeError(
                f"Release record immutable collision: {immutable_path}"
            )
    else:
        immutable_path.write_text(payload, encoding="utf-8")

    latest_path.write_text(payload, encoding="utf-8")
    return immutable_path
