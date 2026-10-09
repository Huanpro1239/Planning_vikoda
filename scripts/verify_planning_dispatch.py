"""Reject spoofed or stale Planning -> NVL dispatch identity before publish."""
from __future__ import annotations

import json
import os
from pathlib import Path

import requests


def validate_release(manifest, *, planning_run_id, head_sha, event, release_id):
    commit = manifest.get("commit") or {}
    if not str(planning_run_id).isdigit():
        raise ValueError("Invalid Planning run_id")
    if len(head_sha) != 40 or any(c not in "0123456789abcdef" for c in head_sha.lower()):
        raise ValueError("Invalid Planning head_sha")
    if event not in {"schedule", "repository_dispatch"}:
        raise ValueError("Invalid Planning event")
    if str(manifest.get("release_id") or "") != release_id:
        raise ValueError("Planning release identity mismatch")
    if str(commit.get("run_id") or "") != str(planning_run_id):
        raise ValueError("Planning release run_id mismatch")
    if str(commit.get("sha") or "") != head_sha:
        raise ValueError("Planning release SHA mismatch")
    if manifest.get("readiness", {}).get("status") != "passed":
        raise ValueError("Planning readiness not passed")
    if manifest.get("publish_decision", {}).get("state") not in {
        "published", "published_no_change"
    }:
        raise ValueError("Planning release was not published")


def main():
    run_id = os.environ["NVL_UPSTREAM_PLANNING_RUN_ID"]
    sha = os.environ["NVL_UPSTREAM_PLANNING_HEAD_SHA"]
    event = os.environ["NVL_UPSTREAM_PLANNING_EVENT"]
    release_id = os.environ["NVL_UPSTREAM_PLANNING_RELEASE_ID"]
    repo = os.environ["GITHUB_REPOSITORY"]
    token = os.environ["GITHUB_TOKEN"]

    # Never trust only the event payload. Require immutable published release.
    release_path = Path(".runtime-state/releases") / f"{release_id}.json"
    if not release_path.is_file():
        raise ValueError("Planning release provenance not persisted: " + release_id)
    manifest = json.loads(release_path.read_text(encoding="utf-8"))
    validate_release(
        manifest,
        planning_run_id=run_id,
        head_sha=sha,
        event=event,
        release_id=release_id,
    )
    response = requests.get(
        f"https://api.github.com/repos/{repo}/actions/runs/{run_id}",
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
        },
        timeout=30,
    )
    response.raise_for_status()
    run = response.json()
    if (
        run.get("name") != "Sync SharePoint Stock"
        or run.get("status") != "completed"
        or run.get("conclusion") != "success"
        or run.get("event") != event
        or run.get("head_sha") != sha
    ):
        raise ValueError("Upstream Planning run identity/status does not match release")
    print(f"[NVL-PAIR] Verified Planning run={run_id} sha={sha} event={event}")


if __name__ == "__main__":
    main()
