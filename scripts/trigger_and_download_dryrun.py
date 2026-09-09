"""Trigger dry-run on sync-nvl-stock.yml, monitor status and conclusion,
download artifacts for diagnostics, and fail with exit code 1 if conclusion != success.

Strictly matches expected commit SHA or run ID to prevent using stale artifacts as evidence.
"""

import argparse
import io
import json
import os
import subprocess
import sys

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import time
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
import requests

REPO = "Huanpro1239/Planning_vikoda"
BRANCH = "feat/sync-nvl-stock"
WORKFLOW = "sync-nvl-stock.yml"


def get_token() -> str | None:
    p = subprocess.Popen(
        ["git", "credential", "fill"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        text=True,
    )
    out, _ = p.communicate("protocol=https\nhost=github.com\n\n")
    creds = dict(line.split("=", 1) for line in out.strip().split("\n") if "=" in line)
    return creds.get("password")


def get_expected_sha(cwd: str | None = None) -> str:
    try:
        out = subprocess.check_output(["git", "rev-parse", "HEAD"], text=True, cwd=cwd).strip()
        return out
    except Exception as exc:
        print(f"Warning: Could not determine git HEAD SHA: {exc}", file=sys.stderr)
        return ""


def dispatch_workflow(token: str, repo: str, workflow: str, branch: str, inputs: dict[str, Any] | None = None) -> None:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    }
    dispatch_url = f"https://api.github.com/repos/{repo}/actions/workflows/{workflow}/dispatches"
    payload = {"ref": branch, "inputs": inputs or {"publish": False}}
    resp = requests.post(dispatch_url, headers=headers, json=payload, timeout=30)
    if resp.status_code not in (200, 204):
        raise RuntimeError(
            f"Workflow dispatch failed with HTTP {resp.status_code}: {resp.text}. "
            f"Không tiếp tục để tránh nhận nhầm run cũ."
        )
    print(f"[TRIGGER] Dispatch workflow '{workflow}' thành công trên ref '{branch}' (HTTP {resp.status_code}).")


def find_run_for_sha(
    token: str,
    repo: str,
    branch: str,
    workflow: str,
    expected_sha: str,
    timeout_seconds: int = 60,
    poll_interval: int = 3,
) -> dict[str, Any]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    }
    runs_url = f"https://api.github.com/repos/{repo}/actions/runs?branch={branch}&per_page=20"
    start_poll = time.time()

    print(f"[TRIGGER] Tìm kiếm workflow run cho SHA '{expected_sha}' trên branch '{branch}' (timeout={timeout_seconds}s)...")
    while time.time() - start_poll < timeout_seconds:
        resp = requests.get(runs_url, headers=headers, timeout=30)
        if resp.status_code == 200:
            data = resp.json()
            for run in data.get("workflow_runs", []):
                head_sha = run.get("head_sha", "")
                head_branch = run.get("head_branch", "")
                run_workflow = run.get("path", "") or run.get("name", "")

                # Kiểm tra khớp branch, workflow và ĐẶC BIỆT là khớp đúng SHA
                is_branch_match = (head_branch == branch)
                is_sha_match = (expected_sha and head_sha == expected_sha) or (not expected_sha)
                is_wf_match = (workflow in run_workflow or run.get("name") == "Sync SharePoint NVL Stock")

                if is_branch_match and is_sha_match and is_wf_match:
                    print(
                        f"[TRIGGER] Đã tìm thấy run {run['id']} ({run.get('name')}): "
                        f"sha={head_sha}, status={run.get('status')}, event={run.get('event')}"
                    )
                    return run
        time.sleep(poll_interval)

    raise TimeoutError(
        f"Hết thời gian ({timeout_seconds}s) chờ workflow run xuất hiện cho SHA '{expected_sha}' trên branch '{branch}'. "
        f"Tuyệt đối không nhận run cũ của SHA khác làm chứng cứ nghiệm thu."
    )


def wait_for_run_completion(
    token: str,
    repo: str,
    run_id: int,
    poll_interval: int = 5,
    max_wait_seconds: int = 600,
) -> dict[str, Any]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    }
    run_url = f"https://api.github.com/repos/{repo}/actions/runs/{run_id}"
    start_wait = time.time()

    print(f"[TRIGGER] Đang theo dõi run {run_id} cho đến khi hoàn thành...")
    while time.time() - start_wait < max_wait_seconds:
        r = requests.get(run_url, headers=headers, timeout=30)
        if r.status_code != 200:
            print(f"[TRIGGER] Warning: GET run {run_id} HTTP {r.status_code}", file=sys.stderr)
            time.sleep(poll_interval)
            continue
        data = r.json()
        status = data.get("status")
        conclusion = data.get("conclusion")
        head_sha = data.get("head_sha")
        print(f"[TRIGGER] Run {run_id}: status={status}, conclusion={conclusion}, sha={head_sha}")
        if status == "completed":
            return data
        time.sleep(poll_interval)

    raise TimeoutError(f"Hết thời gian ({max_wait_seconds}s) chờ run {run_id} hoàn tất.")


def download_run_artifacts(
    token: str,
    repo: str,
    run_id: int,
    target_dir: Path,
    timeout_seconds: int = 30,
) -> list[str]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    }
    art_url = f"https://api.github.com/repos/{repo}/actions/runs/{run_id}/artifacts"
    target_dir.mkdir(parents=True, exist_ok=True)

    arts = []
    start_poll = time.time()
    while time.time() - start_poll < timeout_seconds:
        arts_resp = requests.get(art_url, headers=headers, timeout=30)
        if arts_resp.status_code == 200:
            arts = arts_resp.json().get("artifacts", [])
            if arts:
                break
        time.sleep(2)

    extracted_files: list[str] = []
    if arts:
        for art in arts:
            print(f"[TRIGGER] Đang tải artifact '{art['name']}' ({art['size_in_bytes']} bytes)...")
            dl_resp = requests.get(art["archive_download_url"], headers=headers, timeout=60)
            if dl_resp.status_code == 200:
                with zipfile.ZipFile(io.BytesIO(dl_resp.content)) as z:
                    z.extractall(target_dir)
                    print(f"[TRIGGER] Đã giải nén: {z.namelist()}")
                    extracted_files.extend(z.namelist())
        print(f"[TRIGGER] Toàn bộ artifacts đã lưu vào: {target_dir.resolve()}")
    else:
        print("[TRIGGER] Cảnh báo: Không có artifact nào được tải lên từ run này.")

    return extracted_files


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Trigger và tải artifacts cho NVL dry-run workflow.")
    parser.add_argument("--sha", help="Commit SHA mong đợi (mặc định: git rev-parse HEAD).")
    parser.add_argument("--run-id", type=int, help="Run ID cụ thể cần theo dõi/tải (bỏ qua dispatch).")
    parser.add_argument("--branch", default=BRANCH, help="Tên nhánh GitHub.")
    parser.add_argument("--workflow", default=WORKFLOW, help="Tên file workflow.")
    parser.add_argument("--repo", default=REPO, help="Định dạng owner/repo trên GitHub.")
    parser.add_argument("--no-dispatch", action="store_true", help="Không dispatch mới, chỉ tìm run có sẵn của SHA.")
    parser.add_argument("--out-dir", default="dry_run_artifacts", help="Thư mục lưu artifacts.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    token = get_token()
    if not token:
        print("Lỗi: Không tìm thấy GitHub token qua git credential fill.", file=sys.stderr)
        return 1

    expected_sha = args.sha or get_expected_sha()
    print(f"[TRIGGER] Cấu hình: Repo={args.repo}, Branch={args.branch}, Workflow={args.workflow}")
    print(f"[TRIGGER] Expected commit SHA: {expected_sha}")

    run_id = args.run_id
    head_sha = expected_sha

    if not run_id:
        # Nếu chưa có run-id cụ thể, kiểm tra xem có dispatch mới không
        if not args.no_dispatch:
            try:
                dispatch_workflow(token, args.repo, args.workflow, args.branch, {"publish": False})
            except Exception as exc:
                print(f"[TRIGGER] LỖI DISPATCH: {exc}", file=sys.stderr)
                return 1

        # Tìm run khớp expected_sha
        try:
            run_data = find_run_for_sha(
                token,
                args.repo,
                args.branch,
                args.workflow,
                expected_sha,
                timeout_seconds=60,
            )
            run_id = run_data["id"]
            head_sha = run_data.get("head_sha", expected_sha)
        except Exception as exc:
            print(f"[TRIGGER] LỖI TÌM RUN: {exc}", file=sys.stderr)
            return 1
    else:
        # Nếu chỉ định run_id cụ thể, kiểm tra identity của run
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github.v3+json"}
        r = requests.get(f"https://api.github.com/repos/{args.repo}/actions/runs/{run_id}", headers=headers, timeout=30)
        if r.status_code == 200:
            run_data = r.json()
            run_sha = run_data.get("head_sha", "")
            if expected_sha and run_sha != expected_sha:
                print(
                    f"[TRIGGER] LỖI: Run {run_id} có SHA '{run_sha}' không khớp expected SHA '{expected_sha}'.",
                    file=sys.stderr,
                )
                return 1
            head_sha = run_sha
        else:
            print(f"[TRIGGER] Lỗi truy vấn run {run_id}: HTTP {r.status_code}", file=sys.stderr)
            return 1

    # Chờ hoàn thành
    try:
        completed_run = wait_for_run_completion(token, args.repo, run_id)
    except Exception as exc:
        print(f"[TRIGGER] LỖI CHỜ HOÀN THÀNH: {exc}", file=sys.stderr)
        return 1

    status = completed_run.get("status")
    conclusion = completed_run.get("conclusion")
    head_sha = completed_run.get("head_sha", head_sha)

    # Tải artifacts
    target_dir = Path(args.out_dir)
    download_run_artifacts(token, args.repo, run_id, target_dir)

    print("\n" + "=" * 60)
    print(f"BÁO CÁO KẾT QUẢ WORKFLOW:")
    print(f"  - Workflow:   {args.workflow}")
    print(f"  - Run ID:     {run_id}")
    print(f"  - Commit SHA: {head_sha}")
    print(f"  - Status:     {status}")
    print(f"  - Conclusion: {conclusion}")
    print("=" * 60 + "\n")

    if conclusion != "success":
        print(f"ERROR: Workflow {run_id} kết thúc thất bại với conclusion: {conclusion}", file=sys.stderr)
        return 1

    print(f"SUCCESS: Workflow {run_id} hoàn tất thành công với conclusion: {conclusion}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
