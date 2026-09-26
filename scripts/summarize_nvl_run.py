"""Build the per-run NVL operational summary."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from nvl.ops_summary import (  # noqa: E402
    build_operational_summary,
    write_operational_summary,
)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", default=".")
    parser.add_argument("--job-status", default=os.getenv("JOB_STATUS", "unknown"))
    parser.add_argument("--event", default=os.getenv("GITHUB_EVENT_NAME", ""))
    parser.add_argument("--run-id", default=os.getenv("GITHUB_RUN_ID", ""))
    parser.add_argument(
        "--commit",
        default=(
            os.getenv("NVL_RELEASE_COMMIT_SHA")
            or os.getenv("GITHUB_SHA", "")
        ),
    )
    parser.add_argument("--out", default="nvl_operational_summary.json")
    parser.add_argument(
        "--markdown",
        default=os.getenv("GITHUB_STEP_SUMMARY", ""),
    )
    args = parser.parse_args(argv)

    summary = build_operational_summary(
        args.root,
        job_status=args.job_status,
        event_name=args.event,
        run_id=args.run_id,
        commit_sha=args.commit,
    )
    write_operational_summary(
        summary,
        json_path=args.out,
        markdown_path=args.markdown or None,
    )
    print(
        f"[NVL-OPS] outcome={summary['outcome']} "
        f"phase={summary.get('failed_phase')} "
        f"release={(summary.get('release') or {}).get('release_id')}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
