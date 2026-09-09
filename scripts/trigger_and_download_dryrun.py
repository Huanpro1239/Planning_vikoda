"""Trigger dry-run on sync-nvl-stock.yml, monitor status and conclusion,
download artifacts for diagnostics, and fail with exit code 1 if conclusion != success.
"""

import io
import json
import os
import subprocess
import sys
import time
import zipfile
from pathlib import Path
import requests

REPO = "Huanpro1239/Planning_vikoda"
BRANCH = "feat/sync-nvl-stock"
WORKFLOW = "sync-nvl-stock.yml"


def get_token():
    p = subprocess.Popen(
        ["git", "credential", "fill"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        text=True,
    )
    out, _ = p.communicate("protocol=https\nhost=github.com\n\n")
    creds = dict(line.split("=", 1) for line in out.strip().split("\n") if "=" in line)
    return creds.get("password")


def main():
    token = get_token()
    if not token:
        print("Error: GitHub token not found", file=sys.stderr)
        sys.exit(1)

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    }

    start_time = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(time.time() - 10))
    print(f"Triggering NVL dry-run workflow '{WORKFLOW}' on branch '{BRANCH}' at {start_time}...")

    dispatch_url = f"https://api.github.com/repos/{REPO}/actions/workflows/{WORKFLOW}/dispatches"
    resp = requests.post(dispatch_url, headers=headers, json={"ref": BRANCH, "inputs": {"publish": False}})
    print(f"Dispatch HTTP status: {resp.status_code}")
    if resp.status_code not in (200, 204):
        print(f"Dispatch failed: {resp.text}", file=sys.stderr)
        sys.exit(1)

    run_id = None
    head_sha = None
    runs_url = f"https://api.github.com/repos/{REPO}/actions/workflows/{WORKFLOW}/runs?event=workflow_dispatch&per_page=5"
    for _ in range(30):
        time.sleep(3)
        r = requests.get(runs_url, headers=headers).json()
        for run in r.get("workflow_runs", []):
            if run["created_at"] >= start_time and run["head_branch"] == BRANCH and run["event"] == "workflow_dispatch":
                run_id = run["id"]
                head_sha = run.get("head_sha")
                print(f"Found run {run_id}: status={run['status']}, sha={head_sha}, html={run['html_url']}")
                break
        if run_id:
            break

    if not run_id:
        print("Timed out waiting for run to appear", file=sys.stderr)
        sys.exit(1)

    # Wait for completion
    run_url = f"https://api.github.com/repos/{REPO}/actions/runs/{run_id}"
    conclusion = None
    status = None
    while True:
        r = requests.get(run_url, headers=headers).json()
        status = r.get("status")
        conclusion = r.get("conclusion")
        head_sha = r.get("head_sha", head_sha)
        print(f"Run {run_id}: status={status}, conclusion={conclusion}")
        if status == "completed":
            break
        time.sleep(5)

    # Download artifacts (for diagnostics, even on failure)
    art_url = f"https://api.github.com/repos/{REPO}/actions/runs/{run_id}/artifacts"
    arts = []
    for _ in range(12):
        time.sleep(2)
        arts_resp = requests.get(art_url, headers=headers)
        if arts_resp.status_code == 200:
            arts = arts_resp.json().get("artifacts", [])
            if arts:
                break

    target_dir = Path("dry_run_artifacts")
    target_dir.mkdir(parents=True, exist_ok=True)

    if arts:
        for art in arts:
            print(f"Downloading artifact {art['name']} ({art['size_in_bytes']} bytes)...")
            dl_resp = requests.get(art["archive_download_url"], headers=headers)
            with zipfile.ZipFile(io.BytesIO(dl_resp.content)) as z:
                z.extractall(target_dir)
                print("Extracted files:", z.namelist())
        print(f"Artifacts downloaded and extracted to {target_dir}")
    else:
        print("Warning: No artifacts uploaded by this run.")

    print("\n" + "=" * 50)
    print(f"Workflow Run Summary:")
    print(f"  - Workflow:   {WORKFLOW}")
    print(f"  - Run ID:     {run_id}")
    print(f"  - Commit SHA: {head_sha}")
    print(f"  - Status:     {status}")
    print(f"  - Conclusion: {conclusion}")
    print("=" * 50 + "\n")

    if conclusion != "success":
        print(f"ERROR: Workflow {run_id} failed with conclusion: {conclusion}", file=sys.stderr)
        sys.exit(1)

    print(f"SUCCESS: Workflow {run_id} completed successfully with conclusion: {conclusion}")
    sys.exit(0)


if __name__ == "__main__":
    main()
