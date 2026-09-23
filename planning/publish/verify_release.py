"""Offline verification of immutable Planning release manifests."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
from typing import Any

from .proposal import proposal_output_sha256
from .release import RELEASE_SCHEMA, RELEASE_SCHEMA_VERSION


DEFAULT_EVIDENCE_NAMES = {
    "proposal": "planning_proposal.xlsx",
    "readiness": "production_readiness_report.json",
    "report": "planning_schedule_report.json",
    "decision": "planning_publish_decision.json",
}


class ReleaseVerificationError(RuntimeError):
    """Raised when release evidence violates the manifest contract."""


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise ReleaseVerificationError(
            f"Không đọc được {label}: {path}"
        ) from exc
    if not isinstance(value, dict):
        raise ReleaseVerificationError(
            f"{label} phải là JSON object: {path}"
        )
    return value


def _check(
    checks: list[dict[str, Any]],
    name: str,
    ok: bool,
    *,
    detail: str = "",
    required: bool = True,
) -> None:
    checks.append({
        "name": name,
        "status": "passed" if ok else ("failed" if required else "warning"),
        "required": required,
        "detail": detail,
    })
    if required and not ok:
        raise ReleaseVerificationError(
            f"{name}: {detail or 'verification failed'}"
        )


def _proposal_identity_from_manifest(
    manifest: dict[str, Any],
    stable_output_sha256: str,
) -> str:
    payload = {
        "algorithm": manifest.get("algorithm"),
        "plan_month": manifest.get("plan_month"),
        "input_revision": manifest.get("input_revision") or {},
        "output_sha256": stable_output_sha256,
    }
    canonical = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(canonical).hexdigest()


def _git_commit_exists(
    sha: str,
    *,
    repo_root: Path,
) -> tuple[bool | None, str]:
    if not sha:
        return False, "manifest thiếu commit SHA"
    git_dir = repo_root / ".git"
    if not git_dir.exists():
        return None, "không có .git local để kiểm commit history"

    result = subprocess.run(
        ["git", "cat-file", "-e", f"{sha}^{{commit}}"],
        cwd=repo_root,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    if result.returncode == 0:
        return True, f"commit {sha} tồn tại trong Git history local"
    return False, f"không tìm thấy commit {sha} trong Git history local"


def discover_evidence(
    manifest_path: Path,
) -> dict[str, Path | None]:
    base = manifest_path.parent
    return {
        key: (
            candidate
            if (candidate := base / name).is_file()
            else None
        )
        for key, name in DEFAULT_EVIDENCE_NAMES.items()
    }


def verify_release(
    manifest_path: str | Path,
    *,
    proposal_path: str | Path | None = None,
    readiness_path: str | Path | None = None,
    report_path: str | Path | None = None,
    decision_path: str | Path | None = None,
    strict: bool = False,
    verify_commit: bool = True,
    repo_root: str | Path | None = None,
) -> dict[str, Any]:
    manifest_path = Path(manifest_path)
    if not manifest_path.is_file():
        raise ReleaseVerificationError(
            f"Không tìm thấy release manifest: {manifest_path}"
        )

    manifest = _load_json(manifest_path, "release manifest")
    discovered = discover_evidence(manifest_path)

    proposal = Path(proposal_path) if proposal_path else discovered["proposal"]
    readiness = Path(readiness_path) if readiness_path else discovered["readiness"]
    report_file = Path(report_path) if report_path else discovered["report"]
    decision_file = Path(decision_path) if decision_path else discovered["decision"]

    checks: list[dict[str, Any]] = []
    _check(
        checks,
        "manifest.schema",
        manifest.get("schema") == RELEASE_SCHEMA,
        detail=f"expected={RELEASE_SCHEMA!r}; actual={manifest.get('schema')!r}",
    )
    _check(
        checks,
        "manifest.schema_version",
        manifest.get("schema_version") == RELEASE_SCHEMA_VERSION,
        detail=(
            f"expected={RELEASE_SCHEMA_VERSION}; "
            f"actual={manifest.get('schema_version')!r}"
        ),
    )

    release_id = str(manifest.get("release_id") or "").strip()
    release_version = str(manifest.get("release_version") or "").strip()
    proposal_info = manifest.get("proposal") or {}
    proposal_id = str(proposal_info.get("proposal_id") or "").strip()
    decision = manifest.get("publish_decision") or {}
    commit = manifest.get("commit") or {}
    readiness_info = manifest.get("readiness") or {}

    _check(checks, "manifest.release_id", bool(release_id), detail="release_id rỗng")
    _check(
        checks,
        "manifest.release_version",
        release_version.startswith(f"{RELEASE_SCHEMA}.proposal-"),
        detail=f"release_version không đúng schema: {release_version!r}",
    )
    _check(checks, "manifest.proposal_id", bool(proposal_id), detail="proposal_id rỗng")
    _check(
        checks,
        "manifest.publish_state",
        decision.get("state") in {"published", "published_no_change"},
        detail=f"state={decision.get('state')!r}",
    )
    _check(
        checks,
        "manifest.decision_proposal_id",
        str(decision.get("proposal_id") or "") == proposal_id,
        detail="publish_decision.proposal_id không khớp proposal.proposal_id",
    )
    _check(
        checks,
        "manifest.readiness_status",
        readiness_info.get("status") == "passed",
        detail=f"readiness.status={readiness_info.get('status')!r}",
    )

    commit_sha = str(commit.get("sha") or "").strip()
    readiness_git_sha = str(readiness_info.get("git_sha") or "").strip()
    _check(
        checks,
        "manifest.commit_sha",
        bool(commit_sha),
        detail="commit.sha rỗng",
    )
    if readiness_git_sha:
        _check(
            checks,
            "manifest.commit_matches_readiness",
            readiness_git_sha == commit_sha,
            detail=(
                f"commit.sha={commit_sha}; "
                f"readiness.git_sha={readiness_git_sha}"
            ),
        )

    expected_prefix = (
        f"{manifest.get('plan_month')}-{proposal_id[:12]}-"
    )
    _check(
        checks,
        "manifest.release_id_identity",
        release_id.startswith(expected_prefix),
        detail=(
            f"release_id phải bắt đầu bằng {expected_prefix!r}; "
            f"actual={release_id!r}"
        ),
    )
    _check(
        checks,
        "manifest.release_version_identity",
        f"proposal-{proposal_id[:12]}" in release_version,
        detail="release_version không chứa proposal prefix",
    )

    if verify_commit:
        root = Path(repo_root) if repo_root else Path.cwd()
        exists, detail = _git_commit_exists(commit_sha, repo_root=root)
        if exists is None:
            _check(
                checks,
                "git.commit_exists",
                False,
                detail=detail,
                required=strict,
            )
        else:
            _check(
                checks,
                "git.commit_exists",
                bool(exists),
                detail=detail,
                required=strict,
            )

    evidence = {
        "proposal": str(proposal) if proposal else None,
        "readiness": str(readiness) if readiness else None,
        "report": str(report_file) if report_file else None,
        "decision": str(decision_file) if decision_file else None,
    }

    if proposal and proposal.is_file():
        proposal_bytes = proposal.read_bytes()
        raw_hash = _sha256(proposal_bytes)
        stable_hash = proposal_output_sha256(proposal_bytes)
        _check(
            checks,
            "proposal.raw_sha256",
            raw_hash == proposal_info.get("output_sha256"),
            detail=(
                f"expected={proposal_info.get('output_sha256')}; "
                f"actual={raw_hash}"
            ),
        )
        _check(
            checks,
            "proposal.stable_sha256",
            stable_hash == proposal_info.get("stable_output_sha256"),
            detail=(
                f"expected={proposal_info.get('stable_output_sha256')}; "
                f"actual={stable_hash}"
            ),
        )
        recomputed_id = _proposal_identity_from_manifest(
            manifest,
            stable_hash,
        )
        _check(
            checks,
            "proposal.proposal_id",
            recomputed_id == proposal_id,
            detail=f"expected={proposal_id}; actual={recomputed_id}",
        )
    else:
        _check(
            checks,
            "evidence.proposal",
            False,
            detail="không có planning_proposal.xlsx để recompute artifact hash",
            required=strict,
        )

    if readiness and readiness.is_file():
        readiness_bytes = readiness.read_bytes()
        readiness_report = _load_json(readiness, "readiness report")
        _check(
            checks,
            "readiness.report_sha256",
            _sha256(readiness_bytes) == readiness_info.get("report_sha256"),
            detail="SHA-256 readiness report không khớp manifest",
        )
        _check(
            checks,
            "readiness.status",
            readiness_report.get("status") == "passed",
            detail=f"status={readiness_report.get('status')!r}",
        )
        _check(
            checks,
            "readiness.gate_version",
            readiness_report.get("gate_version") == readiness_info.get("gate_version"),
            detail=(
                f"manifest={readiness_info.get('gate_version')!r}; "
                f"report={readiness_report.get('gate_version')!r}"
            ),
        )
        report_git_sha = str(readiness_report.get("git_sha") or "").strip()
        if report_git_sha:
            _check(
                checks,
                "readiness.git_sha",
                report_git_sha == commit_sha,
                detail=f"manifest={commit_sha}; readiness={report_git_sha}",
            )
    else:
        _check(
            checks,
            "evidence.readiness",
            False,
            detail="không có production_readiness_report.json",
            required=strict,
        )

    if report_file and report_file.is_file():
        audit_report = _load_json(report_file, "planning schedule report")
        _check(
            checks,
            "audit.plan_month",
            audit_report.get("plan_month") == manifest.get("plan_month"),
            detail="plan_month report không khớp manifest",
        )
        _check(
            checks,
            "audit.algorithm",
            audit_report.get("algorithm") == manifest.get("algorithm"),
            detail="algorithm report không khớp manifest",
        )
        _check(
            checks,
            "audit.input_revision",
            (audit_report.get("input_revision") or {})
            == (manifest.get("input_revision") or {}),
            detail="input_revision report không khớp manifest",
        )
        _check(
            checks,
            "audit.proposal_id",
            str(audit_report.get("proposal_id") or "") == proposal_id,
            detail="proposal_id report không khớp manifest",
        )

        report_raw_hash = audit_report.get("output_sha256")
        if report_raw_hash:
            _check(
                checks,
                "audit.output_sha256",
                report_raw_hash == proposal_info.get("output_sha256"),
                detail="output_sha256 report không khớp manifest",
            )

        report_stable_hash = audit_report.get("proposal_output_sha256")
        if report_stable_hash:
            _check(
                checks,
                "audit.stable_output_sha256",
                report_stable_hash == proposal_info.get("stable_output_sha256"),
                detail="proposal_output_sha256 report không khớp manifest",
            )

        pipeline = audit_report.get("pipeline") or {}
        manifest_fingerprints = manifest.get("pipeline_fingerprints") or {}
        for key, expected in manifest_fingerprints.items():
            actual = pipeline.get(key)
            _check(
                checks,
                f"audit.pipeline_fingerprint.{key}",
                actual == expected,
                detail=(
                    f"manifest={expected!r}; report={actual!r}"
                ),
            )
    else:
        _check(
            checks,
            "evidence.report",
            False,
            detail="không có planning_schedule_report.json",
            required=strict,
        )

    if decision_file and decision_file.is_file():
        audit_decision = _load_json(decision_file, "publish decision")
        _check(
            checks,
            "audit.publish_decision",
            audit_decision == decision,
            detail="planning_publish_decision.json không khớp manifest",
        )
    else:
        _check(
            checks,
            "evidence.decision",
            False,
            detail="không có planning_publish_decision.json",
            required=strict,
        )

    warnings = [
        item for item in checks if item["status"] == "warning"
    ]
    return {
        "verified": True,
        "strict": strict,
        "release_id": release_id,
        "release_version": release_version,
        "manifest": str(manifest_path),
        "commit_sha": commit_sha,
        "proposal_id": proposal_id,
        "checks": checks,
        "warnings": warnings,
        "evidence": evidence,
    }
