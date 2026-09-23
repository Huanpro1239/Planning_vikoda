"""Verify historical Planning release manifests and audit evidence."""

from __future__ import annotations

import hashlib
import json
import re
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any

from .proposal import proposal_id, proposal_output_sha256
from .release import (
    SUPPORTED_RELEASE_SCHEMAS,
    canonical_json_sha256,
)


_HEX_64 = re.compile(r"^[0-9a-f]{64}$", re.IGNORECASE)
_HEX_COMMIT = re.compile(r"^[0-9a-f]{40}$", re.IGNORECASE)


def _load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise RuntimeError(f"Không đọc được JSON: {path}") from exc


def _find_repo_root(start: Path | None = None) -> Path | None:
    current = (start or Path.cwd()).resolve()
    for candidate in (current, *current.parents):
        if (candidate / ".git").exists():
            return candidate
    return None


def _infer_evidence(manifest_path: Path) -> dict[str, Path]:
    candidates: dict[str, list[Path]] = {
        "proposal": [manifest_path.parent / "planning_proposal.xlsx"],
        "readiness": [manifest_path.parent / "production_readiness_report.json"],
        "report": [manifest_path.parent / "planning_schedule_report.json"],
        "revision": [manifest_path.parent / "planning_input_revision.json"],
        "decision": [manifest_path.parent / "planning_publish_decision.json"],
    }

    if manifest_path.parent.name == "releases":
        root = manifest_path.parent.parent
        evidence_dir = root / "release_evidence" / manifest_path.stem
        candidates["proposal"].append(evidence_dir / "planning_proposal.xlsx")
        candidates["readiness"].append(evidence_dir / "production_readiness_report.json")
        candidates["report"].append(evidence_dir / "planning_schedule_report.json")
        candidates["revision"].append(evidence_dir / "planning_input_revision.json")
        candidates["decision"].append(evidence_dir / "planning_publish_decision.json")

    if manifest_path.name == "manifest.json":
        candidates["proposal"].insert(0, manifest_path.parent / "planning_proposal.xlsx")
        candidates["readiness"].insert(0, manifest_path.parent / "production_readiness_report.json")
        candidates["report"].insert(0, manifest_path.parent / "planning_schedule_report.json")
        candidates["revision"].insert(0, manifest_path.parent / "planning_input_revision.json")
        candidates["decision"].insert(0, manifest_path.parent / "planning_publish_decision.json")

    result = {}
    for key, paths in candidates.items():
        for path in paths:
            if path.is_file():
                result[key] = path
                break
    return result


def _git_commit_check(repo_root: Path, sha: str) -> tuple[bool, str]:
    exists = subprocess.run(
        ["git", "-C", str(repo_root), "cat-file", "-e", f"{sha}^{{commit}}"],
        check=False,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    if exists.returncode != 0:
        return False, "commit không tồn tại trong local Git object database"

    ancestor = subprocess.run(
        ["git", "-C", str(repo_root), "merge-base", "--is-ancestor", sha, "HEAD"],
        check=False,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    if ancestor.returncode != 0:
        return False, "commit tồn tại nhưng không phải ancestor của HEAD hiện tại"
    return True, "commit tồn tại và là ancestor của HEAD"


def verify_release(
    manifest_path: str | Path,
    *,
    proposal_path: str | Path | None = None,
    readiness_path: str | Path | None = None,
    report_path: str | Path | None = None,
    revision_path: str | Path | None = None,
    decision_path: str | Path | None = None,
    repo_root: str | Path | None = None,
    strict: bool = False,
) -> dict[str, Any]:
    manifest_file = Path(manifest_path)
    if not manifest_file.is_file():
        raise RuntimeError(f"Không tìm thấy manifest: {manifest_file}")

    manifest = _load_json(manifest_file)
    if not isinstance(manifest, dict):
        raise RuntimeError("Release manifest phải là JSON object.")

    inferred = _infer_evidence(manifest_file)
    evidence = {
        "proposal": Path(proposal_path) if proposal_path else inferred.get("proposal"),
        "readiness": Path(readiness_path) if readiness_path else inferred.get("readiness"),
        "report": Path(report_path) if report_path else inferred.get("report"),
        "revision": Path(revision_path) if revision_path else inferred.get("revision"),
        "decision": Path(decision_path) if decision_path else inferred.get("decision"),
    }

    checks: list[dict[str, str]] = []

    def record(name: str, ok: bool, detail: str) -> None:
        checks.append({
            "name": name,
            "status": "passed" if ok else "failed",
            "detail": detail,
        })

    def missing(name: str, detail: str) -> None:
        checks.append({
            "name": name,
            "status": "failed" if strict else "skipped",
            "detail": detail,
        })

    schema = str(manifest.get("schema") or "")
    schema_version = manifest.get("schema_version")
    expected_version = SUPPORTED_RELEASE_SCHEMAS.get(schema)
    effective_version = (
        schema_version
        if isinstance(schema_version, int)
        else 0
    )
    record(
        "manifest_schema",
        expected_version is not None and schema_version == expected_version,
        f"schema={schema!r}; schema_version={schema_version!r}",
    )

    required_top = (
        "release_id",
        "release_version",
        "published_at",
        "plan_month",
        "commit",
        "readiness",
        "proposal",
        "input_revision",
        "publish_decision",
    )
    absent = [key for key in required_top if manifest.get(key) in (None, "")]
    record(
        "manifest_required_fields",
        not absent,
        "đủ trường bắt buộc" if not absent else "thiếu: " + ", ".join(absent),
    )

    try:
        datetime.fromisoformat(str(manifest.get("published_at")))
        record("published_at", True, str(manifest.get("published_at")))
    except Exception:
        record("published_at", False, "published_at không phải ISO-8601 hợp lệ")

    proposal = manifest.get("proposal") or {}
    proposal_id_value = str(proposal.get("proposal_id") or "")
    raw_hash = str(proposal.get("output_sha256") or "")
    stable_hash = str(proposal.get("stable_output_sha256") or "")
    record(
        "proposal_identity_shape",
        bool(_HEX_64.fullmatch(proposal_id_value))
        and bool(_HEX_64.fullmatch(raw_hash))
        and bool(_HEX_64.fullmatch(stable_hash)),
        "proposal_id/raw/stable hash phải là SHA-256",
    )

    release_id = str(manifest.get("release_id") or "")
    release_version = str(manifest.get("release_version") or "")
    plan_month = str(manifest.get("plan_month") or "")
    record(
        "release_version_binding",
        plan_month in release_id
        and proposal_id_value[:12] in release_id
        and schema in release_version
        and proposal_id_value[:12] in release_version,
        f"release_id={release_id}; release_version={release_version}",
    )

    decision = manifest.get("publish_decision") or {}
    record(
        "publish_decision",
        decision.get("state") in {"published", "published_no_change"}
        and str(decision.get("proposal_id") or "") == proposal_id_value,
        f"state={decision.get('state')}; proposal_id={decision.get('proposal_id')}",
    )

    readiness = manifest.get("readiness") or {}
    commit = manifest.get("commit") or {}
    commit_sha = str(commit.get("sha") or "")
    readiness_sha = str(readiness.get("git_sha") or "")
    record(
        "readiness_binding",
        readiness.get("status") == "passed"
        and bool(readiness.get("gate_version"))
        and (not readiness_sha or readiness_sha == commit_sha),
        (
            f"status={readiness.get('status')}; "
            f"gate={readiness.get('gate_version')}; "
            f"git_sha={readiness_sha or '<none>'}"
        ),
    )
    record(
        "commit_sha_shape",
        bool(_HEX_COMMIT.fullmatch(commit_sha)),
        f"commit={commit_sha or '<missing>'}",
    )

    root = Path(repo_root).resolve() if repo_root else _find_repo_root()
    if root is None:
        missing("commit_exists", "không tìm thấy local Git repository; dùng --repo-root")
    elif not _HEX_COMMIT.fullmatch(commit_sha):
        record("commit_exists", False, "commit SHA không hợp lệ")
    else:
        ok, detail = _git_commit_check(root, commit_sha)
        if ok:
            record("commit_exists", True, detail)
        elif strict:
            record("commit_exists", False, detail)
        else:
            checks.append({"name": "commit_exists", "status": "skipped", "detail": detail})

    proposal_file = evidence["proposal"]
    if proposal_file and proposal_file.is_file():
        data = proposal_file.read_bytes()
        computed_id, computed_raw, computed_stable = proposal_id(
            {
                "algorithm": manifest.get("algorithm"),
                "plan_month": manifest.get("plan_month"),
                "input_revision": manifest.get("input_revision") or {},
            },
            data,
        )
        record(
            "proposal_raw_hash",
            computed_raw == raw_hash,
            f"expected={raw_hash}; actual={computed_raw}",
        )
        record(
            "proposal_stable_hash",
            computed_stable == stable_hash,
            f"expected={stable_hash}; actual={computed_stable}",
        )
        record(
            "proposal_id_recomputed",
            computed_id == proposal_id_value,
            f"expected={proposal_id_value}; actual={computed_id}",
        )
    else:
        missing(
            "proposal_bytes",
            "không có proposal workbook; dùng --proposal để recompute hash/proposal_id",
        )

    readiness_file = evidence["readiness"]
    if readiness_file and readiness_file.is_file():
        actual = _load_json(readiness_file)
        actual_raw_hash = hashlib.sha256(readiness_file.read_bytes()).hexdigest()
        record(
            "readiness_report_hash",
            actual_raw_hash == readiness.get("report_sha256"),
            f"expected={readiness.get('report_sha256')}; actual={actual_raw_hash}",
        )
        record(
            "readiness_report_content",
            actual.get("status") == readiness.get("status")
            and actual.get("gate_version") == readiness.get("gate_version")
            and (
                not readiness.get("git_sha")
                or actual.get("git_sha") == readiness.get("git_sha")
            ),
            "readiness status/gate/git_sha khớp manifest",
        )
    else:
        missing("readiness_evidence", "thiếu production_readiness_report.json")

    audit_hashes = manifest.get("audit_evidence") or {}

    report_file = evidence["report"]
    if report_file and report_file.is_file():
        actual = _load_json(report_file)
        report_ok = (
            actual.get("proposal_id") == proposal_id_value
            and actual.get("plan_month") == manifest.get("plan_month")
            and actual.get("algorithm") == manifest.get("algorithm")
            and (actual.get("input_revision") or {}) == (manifest.get("input_revision") or {})
            and (actual.get("publish_decision") or {}) == decision
        )
        record("schedule_report_consistency", report_ok, "schedule report khớp manifest")
        if effective_version >= 2:
            record(
                "schedule_report_hash",
                canonical_json_sha256(actual) == audit_hashes.get("schedule_report_sha256"),
                "canonical schedule report hash",
            )
    else:
        missing("schedule_report_evidence", "thiếu planning_schedule_report.json")

    revision_file = evidence["revision"]
    if revision_file and revision_file.is_file():
        actual = _load_json(revision_file)
        record(
            "input_revision_consistency",
            actual == (manifest.get("input_revision") or {}),
            "input revision khớp manifest",
        )
        if effective_version >= 2:
            record(
                "input_revision_hash",
                canonical_json_sha256(actual) == audit_hashes.get("input_revision_sha256"),
                "canonical input revision hash",
            )
    else:
        missing("input_revision_evidence", "thiếu planning_input_revision.json")

    decision_file = evidence["decision"]
    if decision_file and decision_file.is_file():
        actual = _load_json(decision_file)
        record(
            "publish_decision_consistency",
            actual == decision,
            "publish decision khớp manifest",
        )
        if effective_version >= 2:
            record(
                "publish_decision_hash",
                canonical_json_sha256(actual) == audit_hashes.get("publish_decision_sha256"),
                "canonical publish decision hash",
            )
    else:
        missing("publish_decision_evidence", "thiếu planning_publish_decision.json")

    failures = [item for item in checks if item["status"] == "failed"]
    skipped = [item for item in checks if item["status"] == "skipped"]
    status = "failed" if failures else ("partial" if skipped else "verified")
    return {
        "schema_version": 1,
        "verification_status": status,
        "manifest": str(manifest_file),
        "release_id": manifest.get("release_id"),
        "release_schema": schema,
        "strict": strict,
        "checks": checks,
        "summary": {
            "passed": sum(item["status"] == "passed" for item in checks),
            "failed": len(failures),
            "skipped": len(skipped),
        },
    }
