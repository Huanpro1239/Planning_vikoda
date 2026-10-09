"""Dispatch a paired NVL run explicitly after a published Planning release.

Required because GITHUB_TOKEN-triggered repository_dispatch workflows do not
reliably generate transitive workflow_run events. Only run after the release
has been persisted to runtime-state.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

import requests


def make_payload(manifest: dict, *, run_id: str, head_sha: str, event: str) -> dict:
    commit = manifest.get("commit") or {}
    if event not in {"schedule", "repository_dispatch"}:
        raise ValueError(f"Non-production Planning event: {event}")
    if str(commit.get("run_id")) != str(run_id):
        raise ValueError("Planning release run_id mismatch; dispatch blocked")
    if str(commit.get("sha")) != head_sha:
        raise ValueError("Planning release head_sha mismatch; dispatch blocked")
    if manifest.get("publish_decision", {}).get("state") not in {
        "published", "published_no_change"
    }:
        raise ValueError("Planning release has not been published")
    if manifest.get("readiness", {}).get("status") != "passed":
        raise ValueError("Planning readiness is not PASS")
    release_id = str(manifest.get("release_id") or "").strip()
    if not release_id:
        raise ValueError("Planning release ID is missing")
    return {
        "event_type": "planning_nvl_published",
        "client_payload": {
            "planning_run_id": str(run_id),
            "planning_head_sha": head_sha,
            "planning_event": event,
            "planning_workflow": "Sync SharePoint Stock",
            "planning_release_id": release_id,
        },
    }


def main():
    manifest = json.loads(
        Path("planning_release_manifest.json").read_text(encoding="utf-8")
    )
    run_id = os.environ["GITHUB_RUN_ID"]
    sha = os.environ["GITHUB_SHA"]
    event = os.environ["GITHUB_EVENT_NAME"]
    token = os.environ["GITHUB_TOKEN"]
    repo = os.environ["GITHUB_REPOSITORY"]

    payload = make_payload(manifest, run_id=run_id, head_sha=sha, event=event)

    # This step executes only after runtime-state persist returned success.
    # Verify the release is really present before creating a downstream run.
    release_id = payload["client_payload"]["planning_release_id"]
    persisted = Path(".runtime-state/releases") / f"{release_id}.json"
    if not persisted.is_file():
        raise RuntimeError("Planning release missing from runtime-state: " + release_id)
    persisted_manifest = json.loads(persisted.read_text(encoding="utf-8"))
    if persisted_manifest != manifest:
        raise RuntimeError("Persisted Planning release differs from dispatch release")

    url = f"https://api.github.com/repos/{repo}/dispatches"
    response = requests.post(
        url,
        json=payload,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        },
        timeout=30,
    )
    if response.status_code != 204:
        raise RuntimeError(
            f"Planning-to-NVL dispatch failed HTTP {response.status_code}: "
            + response.text[:300]
        )
    print(
        "[PAIRED-DISPATCH] NVL event sent for Planning "
        f"run={run_id} sha={sha} event={event} release={release_id}"
    )


if __name__ == "__main__":
    main()
