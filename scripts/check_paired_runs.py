"""Check 1:1 health of successful Planning production runs -> NVL."""

from __future__ import annotations

import argparse
from datetime import datetime, timedelta, timezone
import json
import os
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ops.paired_run_health import (  # noqa: E402
    NVL_WORKFLOW_FILE,
    PLANNING_WORKFLOW_FILE,
    evaluate_paired_run_health,
    fetch_workflow_runs,
    load_release_pairs,
    write_health_csv,
    write_health_json,
    write_health_markdown,
)


def _load_runs(path: str) -> list[dict]:
    value = json.loads(Path(path).read_text(encoding="utf-8"))
    if isinstance(value, dict):
        value = value.get("workflow_runs") or []
    if not isinstance(value, list):
        raise ValueError(f"{path} phải chứa JSON array/workflow_runs")
    return [item for item in value if isinstance(item, dict)]


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--repo",
        default=os.getenv("GITHUB_REPOSITORY", "Huanpro1239/Planning_vikoda"),
    )
    parser.add_argument(
        "--runtime-state",
        default=".runtime-state",
    )
    parser.add_argument(
        "--lookback-hours",
        type=int,
        default=168,
    )
    parser.add_argument(
        "--grace-minutes",
        type=int,
        default=20,
    )
    parser.add_argument(
        "--enforce-after",
        required=True,
        help="UTC/offset ISO timestamp where paired-run enforcement begins",
    )
    parser.add_argument("--planning-runs-file")
    parser.add_argument("--nvl-runs-file")
    parser.add_argument("--out-json", default="paired_run_health.json")
    parser.add_argument("--out-csv", default="paired_run_health.csv")
    parser.add_argument("--out-md", default="paired_run_health.md")
    parser.add_argument(
        "--step-summary",
        default=os.getenv("GITHUB_STEP_SUMMARY", ""),
    )
    args = parser.parse_args(argv)

    now = datetime.now(timezone.utc)
    since = now - timedelta(hours=max(1, args.lookback_hours))

    if bool(args.planning_runs_file) != bool(args.nvl_runs_file):
        parser.error(
            "--planning-runs-file và --nvl-runs-file phải dùng cùng nhau"
        )

    if args.planning_runs_file:
        planning_runs = _load_runs(args.planning_runs_file)
        nvl_runs = _load_runs(args.nvl_runs_file)
    else:
        token = os.getenv("GITHUB_TOKEN") or ""
        if not token:
            parser.error("Thiếu GITHUB_TOKEN để đọc GitHub Actions API")
        planning_runs = fetch_workflow_runs(
            token,
            args.repo,
            PLANNING_WORKFLOW_FILE,
            since=since,
        )
        nvl_runs = fetch_workflow_runs(
            token,
            args.repo,
            NVL_WORKFLOW_FILE,
            since=since,
        )

    releases = load_release_pairs(args.runtime_state)
    health = evaluate_paired_run_health(
        planning_runs,
        nvl_runs,
        releases,
        enforce_after=args.enforce_after,
        grace_minutes=args.grace_minutes,
        now=now,
    )

    write_health_json(health, args.out_json)
    write_health_csv(health, args.out_csv)
    write_health_markdown(health, args.out_md)

    if args.step_summary:
        with Path(args.step_summary).open("a", encoding="utf-8") as handle:
            handle.write(Path(args.out_md).read_text(encoding="utf-8"))

    summary = health["summary"]
    print(
        f"[PAIRED-HEALTH] status={health['status']} "
        f"planning={summary['planning_runs_checked']} "
        f"missing={summary['missing_nvl']} "
        f"duplicate={summary['duplicate_nvl']} "
        f"sha_mismatch={summary['head_sha_mismatch']} "
        f"nvl_failed={summary['nvl_downstream_failed']}"
    )
    return 1 if health["status"] == "failed" else 0


if __name__ == "__main__":
    raise SystemExit(main())
