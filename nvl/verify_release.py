"""Offline verification of immutable NVL release manifests."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
from typing import Any

from nvl.release import (
    NVL_RELEASE_SCHEMA,
    NVL_RELEASE_SCHEMA_VERSION,
    stable_workbook_sha256,
)


DEFAULT_EVIDENCE_NAMES = {
    "stock_proposal": "nvl_stock_proposal.xlsx",
    "final_workbook": "nvl_open_po_proposal.xlsx",
    "readiness": "nvl_production_readiness_report.json",
    "stock_report": "nvl_stock_report.json",
    "open_po_report": "nvl_open_po_report.json",
}


class NVLReleaseVerificationError(RuntimeError):
    """Raised when NVL release evidence violates the manifest contract."""


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise NVLReleaseVerificationError(
            f"Không đọc được {label}: {path}"
        ) from exc
    if not isinstance(value, dict):
        raise NVLReleaseVerificationError(
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
    checks.append(
        {
            "name": name,
            "status": "passed" if ok else ("failed" if required else "warning"),
            "required": required,
            "detail": detail,
        }
    )
    if required and not ok:
        raise NVLReleaseVerificationError(
            f"{name}: {detail or 'verification failed'}"
        )


def _git_commit_exists(
    sha: str,
    *,
    repo_root: Path,
) -> tuple[bool | None, str]:
    if not sha:
        return False, "manifest thiếu commit SHA"
    if not (repo_root / ".git").exists():
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


def verify_nvl_release(
    manifest_path: str | Path,
    *,
    stock_proposal_path: str | Path | None = None,
    final_workbook_path: str | Path | None = None,
    readiness_path: str | Path | None = None,
    stock_report_path: str | Path | None = None,
    open_po_report_path: str | Path | None = None,
    strict: bool = False,
    verify_commit: bool = True,
    repo_root: str | Path | None = None,
) -> dict[str, Any]:
    manifest_path = Path(manifest_path)
    if not manifest_path.is_file():
        raise NVLReleaseVerificationError(
            f"Không tìm thấy NVL release manifest: {manifest_path}"
        )

    manifest = _load_json(manifest_path, "NVL release manifest")
    discovered = discover_evidence(manifest_path)

    stock_proposal = (
        Path(stock_proposal_path)
        if stock_proposal_path
        else discovered["stock_proposal"]
    )
    final_workbook = (
        Path(final_workbook_path)
        if final_workbook_path
        else discovered["final_workbook"]
    )
    readiness = (
        Path(readiness_path)
        if readiness_path
        else discovered["readiness"]
    )
    stock_report_file = (
        Path(stock_report_path)
        if stock_report_path
        else discovered["stock_report"]
    )
    open_po_report_file = (
        Path(open_po_report_path)
        if open_po_report_path
        else discovered["open_po_report"]
    )

    checks: list[dict[str, Any]] = []

    _check(
        checks,
        "manifest.schema",
        manifest.get("schema") == NVL_RELEASE_SCHEMA,
        detail=(
            f"expected={NVL_RELEASE_SCHEMA!r}; "
            f"actual={manifest.get('schema')!r}"
        ),
    )
    _check(
        checks,
        "manifest.schema_version",
        manifest.get("schema_version") == NVL_RELEASE_SCHEMA_VERSION,
        detail=(
            f"expected={NVL_RELEASE_SCHEMA_VERSION}; "
            f"actual={manifest.get('schema_version')!r}"
        ),
    )

    release_id = str(manifest.get("release_id") or "").strip()
    release_version = str(manifest.get("release_version") or "").strip()
    chain = manifest.get("chain")
    upstream_planning = manifest.get("upstream_planning")
    commit = manifest.get("commit") or {}
    readiness_info = manifest.get("readiness") or {}
    revisions = manifest.get("input_revision") or {}
    artifacts = manifest.get("artifacts") or {}
    stock_artifact = artifacts.get("stock_proposal") or {}
    final_artifact = artifacts.get("open_po_proposal") or {}
    published = manifest.get("published_workbook") or {}
    evidence = manifest.get("publish_evidence") or {}
    stock_evidence = evidence.get("stock") or {}
    open_evidence = evidence.get("open_po") or {}
    open_nested = open_evidence.get("evidence") or {}

    final_stable = str(published.get("stable_sha256") or "").strip()
    final_proposal_sha = str(published.get("proposal_sha256") or "").strip()
    server_sha = str(published.get("sha256") or "").strip()

    _check(
        checks,
        "manifest.release_id",
        bool(release_id),
        detail="release_id rỗng",
    )
    _check(
        checks,
        "manifest.release_version",
        release_version.startswith(
            f"{NVL_RELEASE_SCHEMA}.workbook-"
        ),
        detail=f"release_version không đúng schema: {release_version!r}",
    )
    if chain is not None:
        _check(
            checks,
            "manifest.chain_shape",
            isinstance(chain, dict),
            detail="chain phải là JSON object khi được khai báo",
        )
        if isinstance(chain, dict):
            previous_release_id = str(
                chain.get("previous_release_id") or ""
            ).strip()
            previous_manifest_sha256 = str(
                chain.get("previous_manifest_sha256") or ""
            ).strip()
            _check(
                checks,
                "manifest.chain_link_pair",
                bool(previous_release_id)
                == bool(previous_manifest_sha256),
                detail=(
                    "previous_release_id và previous_manifest_sha256 "
                    "phải cùng có hoặc cùng rỗng"
                ),
            )
    if upstream_planning is not None:
        _check(
            checks,
            "manifest.upstream_planning_shape",
            isinstance(upstream_planning, dict),
            detail="upstream_planning phải là JSON object khi được khai báo",
        )
        if isinstance(upstream_planning, dict):
            upstream_run_id = str(
                upstream_planning.get("run_id") or ""
            ).strip()
            upstream_head_sha = str(
                upstream_planning.get("head_sha") or ""
            ).strip()
            upstream_event = str(
                upstream_planning.get("event") or ""
            ).strip()
            upstream_workflow = str(
                upstream_planning.get("workflow") or ""
            ).strip()
            upstream_present = any(
                (
                    upstream_run_id,
                    upstream_head_sha,
                    upstream_event,
                    upstream_workflow,
                )
            )
            if upstream_present:
                _check(
                    checks,
                    "manifest.upstream_planning_complete",
                    bool(upstream_run_id)
                    and bool(upstream_head_sha)
                    and bool(upstream_event),
                    detail=(
                        "paired-run provenance phải có "
                        "run_id/head_sha/event"
                    ),
                )
                _check(
                    checks,
                    "manifest.upstream_planning_head_matches_commit",
                    upstream_head_sha
                    == str(commit.get("sha") or "").strip(),
                    detail=(
                        f"upstream.head_sha={upstream_head_sha!r}; "
                        f"commit.sha={commit.get('sha')!r}"
                    ),
                )
    _check(
        checks,
        "manifest.final_stable_sha256",
        bool(final_stable),
        detail="published_workbook.stable_sha256 rỗng",
    )
    _check(
        checks,
        "manifest.final_proposal_sha256",
        bool(final_proposal_sha),
        detail="published_workbook.proposal_sha256 rỗng",
    )
    _check(
        checks,
        "manifest.server_sha256",
        bool(server_sha),
        detail="published_workbook.sha256 rỗng",
    )
    _check(
        checks,
        "manifest.release_id_identity",
        release_id.startswith(
            f"nvl-{final_stable[:16]}-"
        ),
        detail=(
            f"release_id phải bắt đầu bằng "
            f"'nvl-{final_stable[:16]}-'; actual={release_id!r}"
        ),
    )
    _check(
        checks,
        "manifest.release_version_identity",
        f"workbook-{final_stable[:16]}" in release_version,
        detail="release_version không chứa final workbook hash prefix",
    )

    _check(
        checks,
        "manifest.readiness_status",
        readiness_info.get("status") == "passed",
        detail=f"readiness.status={readiness_info.get('status')!r}",
    )
    _check(
        checks,
        "manifest.stock_publish_status",
        stock_evidence.get("status")
        in {"published", "published_with_warnings", "unchanged"},
        detail=f"stock.status={stock_evidence.get('status')!r}",
    )
    _check(
        checks,
        "manifest.stock_post_upload_verified",
        stock_evidence.get("post_upload_verified") is True,
        detail=(
            "publish_evidence.stock.post_upload_verified "
            f"={stock_evidence.get('post_upload_verified')!r}"
        ),
    )
    _check(
        checks,
        "manifest.open_po_published",
        open_evidence.get("published") is True,
        detail=f"open_po.published={open_evidence.get('published')!r}",
    )
    _check(
        checks,
        "manifest.open_po_post_upload_verified",
        open_nested.get("post_upload_verified") is True,
        detail=(
            "publish_evidence.open_po.evidence.post_upload_verified "
            f"={open_nested.get('post_upload_verified')!r}"
        ),
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

    expected_revision_keys = {
        "stock_source_etag",
        "stock_target_etag_before",
        "stock_target_etag_after",
        "open_po_source_etag",
        "open_po_target_etag_before",
        "open_po_target_etag_after",
    }
    _check(
        checks,
        "manifest.input_revision_shape",
        expected_revision_keys.issubset(revisions.keys()),
        detail=(
            "thiếu revision keys: "
            + ", ".join(
                sorted(expected_revision_keys.difference(revisions.keys()))
            )
        ),
    )

    _check(
        checks,
        "manifest.stock_artifact_hashes",
        bool(stock_artifact.get("sha256"))
        and bool(stock_artifact.get("stable_sha256")),
        detail="stock proposal hash chưa đầy đủ",
    )
    _check(
        checks,
        "manifest.final_artifact_hashes",
        bool(final_artifact.get("sha256"))
        and bool(final_artifact.get("stable_sha256")),
        detail="Open-PO proposal hash chưa đầy đủ",
    )
    _check(
        checks,
        "manifest.final_hash_consistency",
        final_artifact.get("sha256") == final_proposal_sha
        and final_artifact.get("stable_sha256") == final_stable,
        detail="published_workbook proposal hashes không khớp artifact",
    )
    _check(
        checks,
        "manifest.server_hash_evidence",
        open_nested.get("server_sha256") == server_sha,
        detail="server SHA trong publish evidence không khớp manifest",
    )

    if verify_commit:
        root = Path(repo_root) if repo_root else Path.cwd()
        exists, detail = _git_commit_exists(
            commit_sha,
            repo_root=root,
        )
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

    evidence_paths = {
        "stock_proposal": (
            str(stock_proposal) if stock_proposal else None
        ),
        "final_workbook": (
            str(final_workbook) if final_workbook else None
        ),
        "readiness": str(readiness) if readiness else None,
        "stock_report": (
            str(stock_report_file) if stock_report_file else None
        ),
        "open_po_report": (
            str(open_po_report_file) if open_po_report_file else None
        ),
    }

    if stock_proposal and stock_proposal.is_file():
        data = stock_proposal.read_bytes()
        raw_hash = _sha256(data)
        stable_hash = stable_workbook_sha256(data)
        _check(
            checks,
            "stock_proposal.raw_sha256",
            raw_hash == stock_artifact.get("sha256"),
            detail=(
                f"expected={stock_artifact.get('sha256')}; "
                f"actual={raw_hash}"
            ),
        )
        _check(
            checks,
            "stock_proposal.stable_sha256",
            stable_hash == stock_artifact.get("stable_sha256"),
            detail=(
                f"expected={stock_artifact.get('stable_sha256')}; "
                f"actual={stable_hash}"
            ),
        )
    else:
        _check(
            checks,
            "evidence.stock_proposal",
            False,
            detail="không có nvl_stock_proposal.xlsx",
            required=strict,
        )

    if final_workbook and final_workbook.is_file():
        data = final_workbook.read_bytes()
        raw_hash = _sha256(data)
        stable_hash = stable_workbook_sha256(data)
        _check(
            checks,
            "final_workbook.proposal_sha256",
            raw_hash == final_artifact.get("sha256")
            and raw_hash == final_proposal_sha,
            detail=(
                f"artifact={final_artifact.get('sha256')}; "
                f"published={final_proposal_sha}; actual={raw_hash}"
            ),
        )
        _check(
            checks,
            "final_workbook.stable_sha256",
            stable_hash == final_artifact.get("stable_sha256")
            and stable_hash == final_stable,
            detail=(
                f"artifact={final_artifact.get('stable_sha256')}; "
                f"published={final_stable}; actual={stable_hash}"
            ),
        )
    else:
        _check(
            checks,
            "evidence.final_workbook",
            False,
            detail="không có nvl_open_po_proposal.xlsx",
            required=strict,
        )

    if readiness and readiness.is_file():
        readiness_bytes = readiness.read_bytes()
        readiness_report = _load_json(
            readiness,
            "NVL readiness report",
        )
        _check(
            checks,
            "readiness.report_sha256",
            _sha256(readiness_bytes)
            == readiness_info.get("report_sha256"),
            detail="SHA-256 NVL readiness report không khớp manifest",
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
            readiness_report.get("gate_version")
            == readiness_info.get("gate_version"),
            detail=(
                f"manifest={readiness_info.get('gate_version')!r}; "
                f"report={readiness_report.get('gate_version')!r}"
            ),
        )
        report_git_sha = str(
            readiness_report.get("git_sha") or ""
        ).strip()
        if report_git_sha:
            _check(
                checks,
                "readiness.git_sha",
                report_git_sha == commit_sha,
                detail=(
                    f"manifest={commit_sha}; "
                    f"readiness={report_git_sha}"
                ),
            )
    else:
        _check(
            checks,
            "evidence.readiness",
            False,
            detail="không có nvl_production_readiness_report.json",
            required=strict,
        )

    if stock_report_file and stock_report_file.is_file():
        stock_report = _load_json(
            stock_report_file,
            "NVL stock report",
        )
        source = stock_report.get("source") or {}
        target = stock_report.get("target") or {}
        upload = stock_report.get("upload_result") or {}
        expected_after = (
            upload.get("eTag")
            or (
                target.get("revision")
                if stock_report.get("status") == "unchanged"
                else None
            )
        )
        _check(
            checks,
            "audit.stock.mode",
            stock_report.get("mode") == "publish",
            detail=f"mode={stock_report.get('mode')!r}",
        )
        _check(
            checks,
            "audit.stock.status",
            stock_report.get("status")
            == stock_evidence.get("status"),
            detail="stock report status không khớp manifest",
        )
        _check(
            checks,
            "audit.stock.source_revision",
            source.get("revision")
            == revisions.get("stock_source_etag"),
            detail="stock source ETag không khớp manifest",
        )
        _check(
            checks,
            "audit.stock.target_revision_before",
            target.get("revision")
            == revisions.get("stock_target_etag_before"),
            detail="stock target ETag trước publish không khớp manifest",
        )
        _check(
            checks,
            "audit.stock.target_revision_after",
            expected_after
            == revisions.get("stock_target_etag_after"),
            detail="stock target ETag sau publish không khớp manifest",
        )
        _check(
            checks,
            "audit.stock.post_upload_verified",
            stock_report.get(
                "post_upload_verified",
                stock_report.get("status") == "unchanged",
            )
            is True,
            detail="stock report không chứng minh post-upload verification",
        )
        _check(
            checks,
            "audit.stock.changed_count",
            (stock_report.get("metrics") or {}).get("changed_count")
            == stock_evidence.get("changed_count"),
            detail="stock changed_count không khớp manifest",
        )
    else:
        _check(
            checks,
            "evidence.stock_report",
            False,
            detail="không có nvl_stock_report.json",
            required=strict,
        )

    if open_po_report_file and open_po_report_file.is_file():
        open_report = _load_json(
            open_po_report_file,
            "NVL Open-PO report",
        )
        source = open_report.get("source") or {}
        target = open_report.get("target") or {}
        report_evidence = open_report.get("publish_evidence") or {}
        _check(
            checks,
            "audit.open_po.publish_requested",
            open_report.get("publish_requested") is True,
            detail="Open-PO report không phải publish run",
        )
        _check(
            checks,
            "audit.open_po.published",
            open_report.get("published") is True,
            detail="Open-PO report chưa published",
        )
        _check(
            checks,
            "audit.open_po.source_revision",
            source.get("revision")
            == revisions.get("open_po_source_etag"),
            detail="Open-PO source ETag không khớp manifest",
        )
        _check(
            checks,
            "audit.open_po.target_revision_before",
            target.get("revision_before")
            == revisions.get("open_po_target_etag_before"),
            detail="Open-PO target ETag trước publish không khớp manifest",
        )
        _check(
            checks,
            "audit.open_po.target_revision_after",
            (
                report_evidence.get("target_revision_after")
                or target.get("revision_after")
            )
            == revisions.get("open_po_target_etag_after"),
            detail="Open-PO target ETag sau publish không khớp manifest",
        )
        _check(
            checks,
            "audit.open_po.publish_evidence",
            report_evidence == open_nested,
            detail="Open-PO publish evidence không khớp manifest",
        )
        _check(
            checks,
            "audit.open_po.changed_cells",
            target.get("changed_cells")
            == open_evidence.get("changed_cells"),
            detail="Open-PO changed_cells không khớp manifest",
        )
        _check(
            checks,
            "audit.open_po.server_sha256",
            report_evidence.get("server_sha256") == server_sha,
            detail="Open-PO server SHA không khớp manifest",
        )
    else:
        _check(
            checks,
            "evidence.open_po_report",
            False,
            detail="không có nvl_open_po_report.json",
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
        "upstream_planning": (
            upstream_planning
            if isinstance(upstream_planning, dict)
            else None
        ),
        "published_workbook_sha256": server_sha,
        "published_workbook_stable_sha256": final_stable,
        "checks": checks,
        "warnings": warnings,
        "evidence": evidence_paths,
    }
