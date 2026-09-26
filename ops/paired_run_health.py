"""Health audit for paired Planning -> NVL production runs."""

from __future__ import annotations

import csv
from datetime import datetime, timedelta, timezone
import json
from pathlib import Path
import re
from typing import Any

import requests


HEALTH_SCHEMA = "paired_run_health_v1"
HEALTH_SCHEMA_VERSION = 1
PLANNING_WORKFLOW_FILE = "sync-stock.yml"
NVL_WORKFLOW_FILE = "sync-nvl-stock.yml"

PAIRED_RUN_NAME_RE = re.compile(
    r"^NVL paired planning=(?P<run_id>\d+) "
    r"sha=(?P<head_sha>[0-9a-fA-F]{7,64}) "
    r"event=(?P<event>schedule|repository_dispatch)$"
)


def _parse_time(value: Any) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def _issue(
    row: dict[str, Any],
    code: str,
    detail: str,
    *,
    severity: str = "error",
) -> None:
    row["issues"].append(
        {
            "severity": severity,
            "code": code,
            "detail": detail,
        }
    )


def parse_paired_nvl_run(run: dict[str, Any]) -> dict[str, Any] | None:
    """Parse upstream Planning identity embedded in the NVL run-name."""
    title = str(run.get("display_title") or run.get("run_name") or "").strip()
    match = PAIRED_RUN_NAME_RE.fullmatch(title)
    if not match:
        return None
    return {
        "planning_run_id": match.group("run_id"),
        "planning_head_sha": match.group("head_sha"),
        "planning_event": match.group("event"),
        "nvl_run_id": str(run.get("id") or ""),
        "nvl_status": run.get("status"),
        "nvl_conclusion": run.get("conclusion"),
        "nvl_created_at": run.get("created_at"),
        "nvl_updated_at": run.get("updated_at"),
        "display_title": title,
        "html_url": run.get("html_url"),
    }


def load_release_pairs(runtime_state: str | Path) -> list[dict[str, Any]]:
    """Load paired provenance directly from immutable NVL manifests."""
    releases_dir = Path(runtime_state) / "nvl" / "releases"
    if not releases_dir.is_dir():
        return []

    records: list[dict[str, Any]] = []
    for path in sorted(releases_dir.glob("*.json")):
        try:
            manifest = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(manifest, dict):
            continue
        upstream = manifest.get("upstream_planning") or {}
        run_id = str(upstream.get("run_id") or "").strip()
        if not run_id:
            continue
        commit = manifest.get("commit") or {}
        records.append(
            {
                "release_id": manifest.get("release_id"),
                "planning_run_id": run_id,
                "planning_head_sha": upstream.get("head_sha"),
                "planning_event": upstream.get("event"),
                "planning_workflow": upstream.get("workflow"),
                "release_commit_sha": commit.get("sha"),
                "published_at": manifest.get("published_at"),
                "path": str(path),
            }
        )
    return records


def _planning_candidates(
    runs: list[dict[str, Any]],
    *,
    enforce_after: datetime,
    now: datetime,
    grace: timedelta,
) -> list[dict[str, Any]]:
    candidates = []
    for run in runs:
        if run.get("event") not in {"schedule", "repository_dispatch"}:
            continue
        if run.get("conclusion") != "success":
            continue

        created_at = _parse_time(run.get("created_at"))
        completed_at = _parse_time(
            run.get("updated_at")
            or run.get("run_started_at")
            or run.get("created_at")
        )
        if created_at is None or completed_at is None:
            continue
        if created_at < enforce_after:
            continue
        if now - completed_at < grace:
            continue
        candidates.append(run)

    candidates.sort(
        key=lambda run: (
            str(run.get("created_at") or ""),
            int(run.get("id") or 0),
        )
    )
    return candidates


def evaluate_paired_run_health(
    planning_runs: list[dict[str, Any]],
    nvl_runs: list[dict[str, Any]],
    release_pairs: list[dict[str, Any]],
    *,
    enforce_after: str | datetime,
    grace_minutes: int = 20,
    now: datetime | None = None,
) -> dict[str, Any]:
    """Evaluate 1:1 Planning -> NVL production pairing."""
    now = (now or datetime.now(timezone.utc)).astimezone(timezone.utc)
    if isinstance(enforce_after, datetime):
        cutoff = enforce_after.astimezone(timezone.utc)
    else:
        cutoff = _parse_time(enforce_after)
        if cutoff is None:
            raise ValueError(f"enforce_after không hợp lệ: {enforce_after!r}")

    grace = timedelta(minutes=max(0, int(grace_minutes)))
    planning = _planning_candidates(
        planning_runs,
        enforce_after=cutoff,
        now=now,
        grace=grace,
    )

    parsed_nvl = [
        pair
        for run in nvl_runs
        if (pair := parse_paired_nvl_run(run)) is not None
    ]

    nvl_by_planning: dict[str, list[dict[str, Any]]] = {}
    for pair in parsed_nvl:
        nvl_by_planning.setdefault(
            pair["planning_run_id"],
            [],
        ).append(pair)

    release_by_planning: dict[str, list[dict[str, Any]]] = {}
    for release in release_pairs:
        run_id = str(release.get("planning_run_id") or "").strip()
        if run_id:
            release_by_planning.setdefault(run_id, []).append(release)

    rows: list[dict[str, Any]] = []

    for plan in planning:
        plan_id = str(plan.get("id") or "")
        plan_sha = str(plan.get("head_sha") or "").strip()
        plan_event = str(plan.get("event") or "").strip()
        downstream = nvl_by_planning.get(plan_id, [])
        releases = release_by_planning.get(plan_id, [])

        row: dict[str, Any] = {
            "planning_run_id": plan_id,
            "planning_head_sha": plan_sha,
            "planning_event": plan_event,
            "planning_created_at": plan.get("created_at"),
            "planning_updated_at": plan.get("updated_at"),
            "planning_url": plan.get("html_url"),
            "nvl_run_count": len(downstream),
            "nvl_run_ids": [item.get("nvl_run_id") for item in downstream],
            "nvl_conclusions": [
                item.get("nvl_conclusion") for item in downstream
            ],
            "release_count": len(releases),
            "release_ids": [
                item.get("release_id") for item in releases
            ],
            "status": "unknown",
            "issues": [],
        }

        if not downstream:
            _issue(
                row,
                "missing_nvl",
                (
                    f"Planning run {plan_id} thành công nhưng không tìm thấy "
                    "NVL downstream workflow_run."
                ),
            )
        elif len(downstream) > 1:
            _issue(
                row,
                "duplicate_nvl",
                (
                    f"Planning run {plan_id} có {len(downstream)} "
                    "NVL downstream runs."
                ),
            )

        for pair in downstream:
            if pair.get("planning_head_sha") != plan_sha:
                _issue(
                    row,
                    "head_sha_mismatch",
                    (
                        f"Planning head_sha={plan_sha}; "
                        f"NVL run {pair.get('nvl_run_id')} encoded "
                        f"head_sha={pair.get('planning_head_sha')}"
                    ),
                )
            if pair.get("planning_event") != plan_event:
                _issue(
                    row,
                    "trigger_event_mismatch",
                    (
                        f"Planning event={plan_event}; "
                        f"NVL encoded event={pair.get('planning_event')}"
                    ),
                )

            status = str(pair.get("nvl_status") or "")
            conclusion = pair.get("nvl_conclusion")
            if status == "completed" and conclusion != "success":
                _issue(
                    row,
                    "nvl_downstream_failed",
                    (
                        f"NVL run {pair.get('nvl_run_id')} "
                        f"conclusion={conclusion!r}"
                    ),
                )
            elif status != "completed":
                _issue(
                    row,
                    "nvl_downstream_incomplete",
                    (
                        f"NVL run {pair.get('nvl_run_id')} "
                        f"status={status!r}"
                    ),
                )

        if len(releases) > 1:
            _issue(
                row,
                "duplicate_release_provenance",
                (
                    f"Planning run {plan_id} xuất hiện trong "
                    f"{len(releases)} NVL release manifests."
                ),
            )

        for release in releases:
            upstream_sha = str(
                release.get("planning_head_sha") or ""
            ).strip()
            commit_sha = str(
                release.get("release_commit_sha") or ""
            ).strip()
            release_event = str(
                release.get("planning_event") or ""
            ).strip()

            if upstream_sha != plan_sha:
                _issue(
                    row,
                    "release_head_sha_mismatch",
                    (
                        f"Planning head_sha={plan_sha}; "
                        f"release {release.get('release_id')} upstream "
                        f"head_sha={upstream_sha}"
                    ),
                )
            if commit_sha != plan_sha:
                _issue(
                    row,
                    "release_commit_sha_mismatch",
                    (
                        f"Planning head_sha={plan_sha}; "
                        f"release {release.get('release_id')} commit "
                        f"sha={commit_sha}"
                    ),
                )
            if release_event != plan_event:
                _issue(
                    row,
                    "release_event_mismatch",
                    (
                        f"Planning event={plan_event}; "
                        f"release event={release_event}"
                    ),
                )

        successful_downstream = any(
            item.get("nvl_status") == "completed"
            and item.get("nvl_conclusion") == "success"
            for item in downstream
        )
        if successful_downstream and not releases:
            _issue(
                row,
                "missing_release_provenance",
                (
                    f"NVL downstream của Planning run {plan_id} success "
                    "nhưng runtime-state không có paired release manifest."
                ),
            )

        row["status"] = (
            "failed"
            if any(
                issue.get("severity") == "error"
                for issue in row["issues"]
            )
            else "passed"
        )
        rows.append(row)

    error_counts: dict[str, int] = {}
    for row in rows:
        for issue in row["issues"]:
            code = str(issue.get("code") or "")
            error_counts[code] = error_counts.get(code, 0) + 1

    status = (
        "empty"
        if not rows
        else (
            "failed"
            if any(row["status"] == "failed" for row in rows)
            else "passed"
        )
    )

    return {
        "schema": HEALTH_SCHEMA,
        "schema_version": HEALTH_SCHEMA_VERSION,
        "generated_at": now.isoformat(),
        "enforce_after": cutoff.isoformat(),
        "grace_minutes": int(grace.total_seconds() // 60),
        "status": status,
        "summary": {
            "planning_runs_checked": len(rows),
            "passed_pairs": sum(
                1 for row in rows if row["status"] == "passed"
            ),
            "failed_pairs": sum(
                1 for row in rows if row["status"] == "failed"
            ),
            "missing_nvl": error_counts.get("missing_nvl", 0),
            "duplicate_nvl": error_counts.get("duplicate_nvl", 0),
            "head_sha_mismatch": (
                error_counts.get("head_sha_mismatch", 0)
                + error_counts.get("release_head_sha_mismatch", 0)
                + error_counts.get("release_commit_sha_mismatch", 0)
            ),
            "nvl_downstream_failed": error_counts.get(
                "nvl_downstream_failed",
                0,
            ),
            "missing_release_provenance": error_counts.get(
                "missing_release_provenance",
                0,
            ),
        },
        "pairs": rows,
    }


def fetch_workflow_runs(
    token: str,
    repo: str,
    workflow_file: str,
    *,
    since: datetime,
    max_pages: int = 10,
) -> list[dict[str, Any]]:
    """Fetch workflow runs from GitHub Actions within a time window."""
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    collected: list[dict[str, Any]] = []

    for page in range(1, max_pages + 1):
        url = (
            f"https://api.github.com/repos/{repo}/actions/workflows/"
            f"{workflow_file}/runs"
        )
        response = requests.get(
            url,
            headers=headers,
            params={"per_page": 100, "page": page},
            timeout=30,
        )
        if response.status_code != 200:
            raise RuntimeError(
                f"GitHub Actions API {workflow_file} HTTP "
                f"{response.status_code}: {response.text}"
            )

        batch = response.json().get("workflow_runs") or []
        if not batch:
            break

        stop = False
        for run in batch:
            created = _parse_time(run.get("created_at"))
            if created is not None and created < since:
                stop = True
                continue
            collected.append(run)

        if stop or len(batch) < 100:
            break

    return collected


def write_health_json(
    health: dict[str, Any],
    path: str | Path,
) -> Path:
    path = Path(path)
    path.write_text(
        json.dumps(
            health,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
    )
    return path


def write_health_csv(
    health: dict[str, Any],
    path: str | Path,
) -> Path:
    path = Path(path)
    fields = [
        "planning_run_id",
        "planning_head_sha",
        "planning_event",
        "planning_created_at",
        "planning_updated_at",
        "nvl_run_count",
        "nvl_run_ids",
        "nvl_conclusions",
        "release_count",
        "release_ids",
        "status",
        "issues",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for row in health.get("pairs", []):
            writer.writerow(
                {
                    field: (
                        " | ".join(
                            f"{issue.get('code')}: {issue.get('detail')}"
                            for issue in row.get("issues", [])
                        )
                        if field == "issues"
                        else (
                            ",".join(str(v) for v in row.get(field, []))
                            if isinstance(row.get(field), list)
                            else row.get(field)
                        )
                    )
                    for field in fields
                }
            )
    return path


def render_health_markdown(health: dict[str, Any]) -> str:
    summary = health.get("summary") or {}
    lines = [
        "# Planning → NVL Paired-Run Health",
        "",
        f"**Status:** {health.get('status')}",
        "",
        "| Signal | Count |",
        "|---|---:|",
        f"| Planning runs checked | {summary.get('planning_runs_checked', 0)} |",
        f"| Passed pairs | {summary.get('passed_pairs', 0)} |",
        f"| Failed pairs | {summary.get('failed_pairs', 0)} |",
        f"| Missing NVL | {summary.get('missing_nvl', 0)} |",
        f"| Duplicate NVL | {summary.get('duplicate_nvl', 0)} |",
        f"| Head SHA mismatch | {summary.get('head_sha_mismatch', 0)} |",
        f"| NVL downstream failed | {summary.get('nvl_downstream_failed', 0)} |",
        (
            f"| Missing release provenance | "
            f"{summary.get('missing_release_provenance', 0)} |"
        ),
        "",
        "## Pair details",
        "",
        "| Planning run | Event | NVL runs | Releases | Status |",
        "|---:|---|---|---|---|",
    ]
    for row in health.get("pairs", []):
        lines.append(
            f"| {row.get('planning_run_id')} | "
            f"{row.get('planning_event')} | "
            f"{','.join(str(v) for v in row.get('nvl_run_ids', [])) or '-'} | "
            f"{','.join(str(v) for v in row.get('release_ids', [])) or '-'} | "
            f"{row.get('status')} |"
        )
        for issue in row.get("issues", []):
            lines.append(
                f"|  |  |  | {issue.get('code')} | "
                f"{issue.get('detail')} |"
            )
    return "\n".join(lines) + "\n"


def write_health_markdown(
    health: dict[str, Any],
    path: str | Path,
) -> Path:
    path = Path(path)
    path.write_text(
        render_health_markdown(health),
        encoding="utf-8",
    )
    return path
