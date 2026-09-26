"""Unified read-only overview for Planning and NVL release history."""

from __future__ import annotations

import csv
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
from typing import Any

from nvl.release_audit import audit_nvl_release_ledger
from planning.publish.verify_release import (
    ReleaseVerificationError,
    verify_release,
)


OVERVIEW_SCHEMA = "release_overview_v1"
OVERVIEW_SCHEMA_VERSION = 1


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _load_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"{path} phải là JSON object")
    return value


def _parse_time(value: Any) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)
    except ValueError:
        return None


def _planning_ledger(
    root: Path,
    *,
    verify_commit: bool,
    repo_root: Path,
) -> dict[str, Any]:
    releases_dir = root / "releases"
    latest_path = root / "latest_release.json"
    paths = sorted(releases_dir.glob("*.json")) if releases_dir.is_dir() else []

    records: list[dict[str, Any]] = []
    errors = 0
    warnings = 0

    for path in paths:
        record: dict[str, Any] = {
            "domain": "planning",
            "path": str(path),
            "filename": path.name,
            "manifest_sha256": _sha256(path.read_bytes()),
            "release_id": None,
            "published_at": None,
            "commit_sha": None,
            "gate_version": None,
            "plan_month": None,
            "proposal_id": None,
            "artifact_sha256": None,
            "status": "unknown",
            "issues": [],
            "_published_dt": None,
        }

        try:
            manifest = _load_json(path)
        except Exception as exc:
            record["status"] = "failed"
            record["issues"].append(
                {
                    "severity": "error",
                    "code": "invalid_manifest",
                    "detail": str(exc),
                }
            )
            errors += 1
            records.append(record)
            continue

        release_id = str(manifest.get("release_id") or "").strip()
        published_at = str(manifest.get("published_at") or "").strip()
        commit = manifest.get("commit") or {}
        readiness = manifest.get("readiness") or {}
        proposal = manifest.get("proposal") or {}

        record.update(
            {
                "release_id": release_id or None,
                "published_at": published_at or None,
                "commit_sha": commit.get("sha"),
                "gate_version": readiness.get("gate_version"),
                "plan_month": manifest.get("plan_month"),
                "proposal_id": proposal.get("proposal_id"),
                "artifact_sha256": proposal.get("output_sha256"),
                "_published_dt": _parse_time(published_at),
            }
        )

        if path.name != f"{release_id}.json":
            record["issues"].append(
                {
                    "severity": "error",
                    "code": "filename_release_id_mismatch",
                    "detail": (
                        f"filename={path.name!r}; "
                        f"expected={release_id + '.json'!r}"
                    ),
                }
            )

        if record["_published_dt"] is None:
            record["issues"].append(
                {
                    "severity": "error",
                    "code": "invalid_published_at",
                    "detail": f"published_at={published_at!r}",
                }
            )

        try:
            result = verify_release(
                path,
                strict=False,
                verify_commit=verify_commit,
                repo_root=repo_root,
            )
            verification_warnings = result.get("warnings") or []
            for warning in verification_warnings:
                record["issues"].append(
                    {
                        "severity": "warning",
                        "code": warning.get("name", "verification_warning"),
                        "detail": warning.get("detail", ""),
                    }
                )
        except ReleaseVerificationError as exc:
            record["issues"].append(
                {
                    "severity": "error",
                    "code": "manifest_verification_failed",
                    "detail": str(exc),
                }
            )

        has_error = any(
            issue["severity"] == "error"
            for issue in record["issues"]
        )
        has_warning = any(
            issue["severity"] == "warning"
            for issue in record["issues"]
        )
        record["status"] = (
            "failed" if has_error else ("warning" if has_warning else "passed")
        )
        errors += sum(
            1 for issue in record["issues"] if issue["severity"] == "error"
        )
        warnings += sum(
            1 for issue in record["issues"] if issue["severity"] == "warning"
        )
        records.append(record)

    valid = [
        row
        for row in records
        if row.get("_published_dt") is not None and row.get("release_id")
    ]
    valid.sort(key=lambda row: (row["_published_dt"], row["release_id"]))

    latest = {
        "path": str(latest_path),
        "status": "not_applicable" if not records else "unknown",
        "release_id": None,
        "detail": "",
    }
    if records:
        newest = valid[-1] if valid else None
        if not latest_path.is_file():
            latest.update(
                {"status": "failed", "detail": "Thiếu latest_release.json"}
            )
            errors += 1
        else:
            try:
                latest_raw = latest_path.read_bytes()
                latest_manifest = json.loads(latest_raw.decode("utf-8"))
                latest_id = str(
                    latest_manifest.get("release_id") or ""
                ).strip()
                latest["release_id"] = latest_id or None
                if newest is None or latest_id != newest.get("release_id"):
                    latest.update(
                        {
                            "status": "failed",
                            "detail": (
                                f"latest={latest_id!r}; "
                                f"newest={newest.get('release_id') if newest else None!r}"
                            ),
                        }
                    )
                    errors += 1
                elif _sha256(latest_raw) != newest["manifest_sha256"]:
                    latest.update(
                        {
                            "status": "failed",
                            "detail": (
                                "latest_release.json không byte-identical "
                                "với newest immutable release"
                            ),
                        }
                    )
                    errors += 1
                else:
                    latest.update(
                        {
                            "status": "passed",
                            "detail": "latest pointer khớp newest Planning release",
                        }
                    )
            except Exception as exc:
                latest.update(
                    {"status": "failed", "detail": str(exc)}
                )
                errors += 1
    elif latest_path.exists():
        latest.update(
            {
                "status": "failed",
                "detail": "Có latest_release.json nhưng không có releases/*.json",
            }
        )
        errors += 1

    for sequence, row in enumerate(valid, start=1):
        row["sequence"] = sequence
    for row in records:
        row.pop("_published_dt", None)
        row.setdefault("sequence", None)

    status = (
        "empty"
        if not records and not latest_path.exists()
        else ("failed" if errors else "passed")
    )
    return {
        "status": status,
        "latest": latest,
        "summary": {
            "release_count": len(records),
            "failed_count": sum(
                1 for row in records if row["status"] == "failed"
            ),
            "warning_count": sum(
                1 for row in records if row["status"] == "warning"
            ),
            "issue_error_count": errors,
            "issue_warning_count": warnings,
        },
        "releases": sorted(
            records,
            key=lambda row: (
                str(row.get("published_at") or ""),
                str(row.get("release_id") or ""),
            ),
        ),
    }


def _nvl_rows(index: dict[str, Any]) -> list[dict[str, Any]]:
    rows = []
    for item in index.get("releases", []):
        rows.append(
            {
                "domain": "nvl",
                "sequence": item.get("sequence"),
                "published_at": item.get("published_at"),
                "release_id": item.get("release_id"),
                "commit_sha": item.get("commit_sha"),
                "gate_version": item.get("gate_version"),
                "plan_month": None,
                "proposal_id": None,
                "artifact_sha256": item.get("published_workbook_sha256"),
                "upstream_planning_run_id": item.get(
                    "upstream_planning_run_id"
                ),
                "upstream_planning_head_sha": item.get(
                    "upstream_planning_head_sha"
                ),
                "upstream_planning_event": item.get(
                    "upstream_planning_event"
                ),
                "status": item.get("status"),
                "chain_status": item.get("chain_status"),
                "issues": item.get("issues") or [],
            }
        )
    return rows


def build_release_overview(
    runtime_state: str | Path,
    *,
    verify_commit: bool = True,
    repo_root: str | Path | None = None,
) -> dict[str, Any]:
    root = Path(runtime_state)
    repo_root_path = Path(repo_root) if repo_root else Path.cwd()

    planning = _planning_ledger(
        root,
        verify_commit=verify_commit,
        repo_root=repo_root_path,
    )
    nvl = audit_nvl_release_ledger(
        root / "nvl" / "releases",
        latest_path=root / "nvl" / "latest_release.json",
        verify_commit=verify_commit,
        repo_root=repo_root_path,
    )

    planning_rows = planning["releases"]
    nvl_rows = _nvl_rows(nvl)
    timeline = planning_rows + nvl_rows
    timeline.sort(
        key=lambda row: (
            str(row.get("published_at") or ""),
            row.get("domain") or "",
            row.get("release_id") or "",
        )
    )
    for sequence, row in enumerate(timeline, start=1):
        row["global_sequence"] = sequence

    statuses = {planning["status"], nvl["status"]}
    overall = (
        "failed"
        if "failed" in statuses
        else (
            "empty"
            if statuses <= {"empty"}
            else "passed"
        )
    )

    def latest_row(domain: str) -> dict[str, Any] | None:
        candidates = [
            row for row in timeline if row.get("domain") == domain
        ]
        return candidates[-1] if candidates else None

    return {
        "schema": OVERVIEW_SCHEMA,
        "schema_version": OVERVIEW_SCHEMA_VERSION,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "status": overall,
        "runtime_state": str(root),
        "planning": planning,
        "nvl": nvl,
        "latest": {
            "planning": latest_row("planning"),
            "nvl": latest_row("nvl"),
        },
        "summary": {
            "planning_release_count": planning["summary"]["release_count"],
            "planning_status": planning["status"],
            "nvl_release_count": nvl["summary"]["release_count"],
            "nvl_status": nvl["status"],
            "total_release_count": len(timeline),
        },
        "timeline": timeline,
    }


def write_overview_json(
    overview: dict[str, Any],
    path: str | Path,
) -> Path:
    path = Path(path)
    path.write_text(
        json.dumps(overview, ensure_ascii=False, indent=2, sort_keys=True)
        + "\n",
        encoding="utf-8",
    )
    return path


def write_overview_csv(
    overview: dict[str, Any],
    path: str | Path,
) -> Path:
    path = Path(path)
    fields = [
        "global_sequence",
        "domain",
        "published_at",
        "release_id",
        "commit_sha",
        "gate_version",
        "plan_month",
        "proposal_id",
        "artifact_sha256",
        "upstream_planning_run_id",
        "upstream_planning_head_sha",
        "upstream_planning_event",
        "status",
        "chain_status",
        "issues",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for row in overview.get("timeline", []):
            writer.writerow(
                {
                    field: (
                        " | ".join(
                            f"{issue.get('code')}: {issue.get('detail')}"
                            for issue in row.get("issues", [])
                        )
                        if field == "issues"
                        else row.get(field)
                    )
                    for field in fields
                }
            )
    return path


def render_overview_markdown(overview: dict[str, Any]) -> str:
    summary = overview.get("summary") or {}
    latest = overview.get("latest") or {}
    planning = latest.get("planning") or {}
    nvl = latest.get("nvl") or {}

    lines = [
        "# Planning + NVL Release Overview",
        "",
        f"**Overall:** {overview.get('status')}",
        "",
        "| Domain | Status | Releases | Latest release | Commit | Artifact hash |",
        "|---|---|---:|---|---|---|",
        (
            f"| Planning | {summary.get('planning_status')} | "
            f"{summary.get('planning_release_count')} | "
            f"{planning.get('release_id') or '-'} | "
            f"{str(planning.get('commit_sha') or '-')[:12]} | "
            f"{str(planning.get('artifact_sha256') or '-')[:16]} |"
        ),
        (
            f"| NVL | {summary.get('nvl_status')} | "
            f"{summary.get('nvl_release_count')} | "
            f"{nvl.get('release_id') or '-'} | "
            f"{str(nvl.get('commit_sha') or '-')[:12]} | "
            f"{str(nvl.get('artifact_sha256') or '-')[:16]} |"
        ),
        "",
        "## Combined timeline",
        "",
        "| # | Time | Domain | Release | Status |",
        "|---:|---|---|---|---|",
    ]
    for row in overview.get("timeline", []):
        lines.append(
            f"| {row.get('global_sequence')} | "
            f"{row.get('published_at') or '-'} | "
            f"{row.get('domain')} | "
            f"{row.get('release_id') or '-'} | "
            f"{row.get('status')} |"
        )
    return "\n".join(lines) + "\n"


def write_overview_markdown(
    overview: dict[str, Any],
    path: str | Path,
) -> Path:
    path = Path(path)
    path.write_text(render_overview_markdown(overview), encoding="utf-8")
    return path
