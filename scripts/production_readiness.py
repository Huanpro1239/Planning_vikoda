"""Single production-readiness gate for Planning releases.

Usage:
    python -X utf8 scripts/production_readiness.py

The command is intentionally identical in PR CI and the production publish
workflow. A release is ready only when every phase passes.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import sys
import time
from datetime import datetime, timezone


ROOT = Path(__file__).resolve().parents[1]
REPORT_PATH = ROOT / "production_readiness_report.json"
GATE_VERSION = "production_readiness_v1"


def _command(*args: str) -> list[str]:
    return [sys.executable, *args]


PHASES = (
    (
        "compile",
        _command("-m", "compileall", "-q", "."),
    ),
    (
        "architecture",
        _command(
            "-m",
            "unittest",
            "-v",
            "tests.test_planning_package_architecture",
            "tests.test_repo_hygiene",
        ),
    ),
    (
        "business_contracts",
        _command(
            "-m",
            "unittest",
            "discover",
            "-s",
            "tests/contracts",
            "-t",
            ".",
            "-p",
            "test_*.py",
            "-v",
        ),
    ),
    (
        "release_safety",
        _command(
            "-m",
            "unittest",
            "discover",
            "-s",
            "tests/readiness",
            "-t",
            ".",
            "-p",
            "test_*.py",
            "-v",
        ),
    ),
    (
        "full_regression",
        _command(
            "-m",
            "unittest",
            "discover",
            "-s",
            "tests",
            "-t",
            ".",
            "-p",
            "test_*.py",
            "-v",
        ),
    ),
)


def _write_report(report: dict) -> None:
    REPORT_PATH.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def main() -> int:
    report = {
        "schema_version": 1,
        "gate_version": GATE_VERSION,
        "status": "running",
        "started_at": datetime.now(timezone.utc).isoformat(),
        "git_sha": os.getenv("GITHUB_SHA") or "",
        "python": sys.version,
        "phases": [],
    }
    _write_report(report)

    print(f"[READINESS] gate={GATE_VERSION}")
    for name, command in PHASES:
        started = time.monotonic()
        print(f"[READINESS][{name}] START")
        result = subprocess.run(
            command,
            cwd=ROOT,
            check=False,
        )
        duration = round(time.monotonic() - started, 3)
        phase = {
            "name": name,
            "status": "passed" if result.returncode == 0 else "failed",
            "returncode": result.returncode,
            "duration_seconds": duration,
            "command": command,
        }
        report["phases"].append(phase)
        _write_report(report)

        if result.returncode != 0:
            report["status"] = "failed"
            report["failed_phase"] = name
            report["finished_at"] = datetime.now(timezone.utc).isoformat()
            _write_report(report)
            print(
                f"[READINESS][{name}] FAIL rc={result.returncode}; "
                f"report={REPORT_PATH.name}"
            )
            return result.returncode

        print(f"[READINESS][{name}] PASS ({duration:.3f}s)")

    report["status"] = "passed"
    report["finished_at"] = datetime.now(timezone.utc).isoformat()
    _write_report(report)
    print(f"[READINESS] PASS report={REPORT_PATH.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
