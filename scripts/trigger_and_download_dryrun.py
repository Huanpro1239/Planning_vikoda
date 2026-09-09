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


def get_existing_run_ids(
    token: str,
    repo: str,
    branch: str,
    workflow: str,
) -> set[int]:
    """Lấy danh sách các run ID hiện có trên GitHub trước khi dispatch để tránh nhận nhầm run cũ."""
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    }
    runs_url = f"https://api.github.com/repos/{repo}/actions/runs?branch={branch}&per_page=50"
    resp = requests.get(runs_url, headers=headers, timeout=30)
    if resp.status_code == 200:
        data = resp.json()
        return {r["id"] for r in data.get("workflow_runs", [])}
    return set()


def find_run_for_sha(
    token: str,
    repo: str,
    branch: str,
    workflow: str,
    expected_sha: str,
    *,
    exclude_run_ids: set[int] | None = None,
    require_event: str | None = None,
    timeout_seconds: int = 60,
    poll_interval: int = 3,
) -> dict[str, Any]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    }
    runs_url = f"https://api.github.com/repos/{repo}/actions/runs?branch={branch}&per_page=20"
    start_poll = time.time()
    excluded = exclude_run_ids or set()

    print(
        f"[TRIGGER] Tìm kiếm workflow run cho SHA '{expected_sha}' trên branch '{branch}' "
        f"(require_event={require_event}, exclude={len(excluded)} runs cũ, timeout={timeout_seconds}s)..."
    )
    while time.time() - start_poll < timeout_seconds:
        resp = requests.get(runs_url, headers=headers, timeout=30)
        if resp.status_code == 200:
            data = resp.json()
            for run in data.get("workflow_runs", []):
                rid = run.get("id")
                # Bỏ qua các run ID đã tồn tại trước lần dispatch này
                if rid in excluded:
                    continue

                run_event = run.get("event", "")
                # Bắt buộc khớp event nếu có yêu cầu (ví dụ: workflow_dispatch)
                if require_event and run_event != require_event:
                    continue

                head_sha = run.get("head_sha", "")
                head_branch = run.get("head_branch", "")
                run_workflow = run.get("path", "") or run.get("name", "")

                # Kiểm tra khớp branch, workflow và ĐẶC BIỆT là khớp đúng SHA
                is_branch_match = (head_branch == branch)
                is_sha_match = (expected_sha and (head_sha == expected_sha or head_sha.startswith(expected_sha))) or (not expected_sha)
                is_wf_match = (workflow in run_workflow or run.get("name") == "Sync SharePoint NVL Stock")

                if is_branch_match and is_sha_match and is_wf_match:
                    print(
                        f"[TRIGGER] Đã tìm thấy run {rid} ({run.get('name')}): "
                        f"sha={head_sha}, status={run.get('status')}, event={run_event}"
                    )
                    return run
        time.sleep(poll_interval)

    raise TimeoutError(
        f"Hết thời gian ({timeout_seconds}s) chờ workflow run mới xuất hiện cho SHA '{expected_sha}' trên branch '{branch}'."
    )


def verify_run_execution(
    artifacts_dir: Path,
    expected_config_file: str,
    expected_publish: bool,
) -> None:
    """Xác minh artifact sau khi tải:
    1. Publish mode: Nếu yêu cầu publish=False thì report không được là mode='publish'.
       Nếu yêu cầu publish=True thì report bắt buộc là mode='publish'.
    2. Cấu hình: Kiểm tra sharepoint_path và target_name trong report/audit phải khớp với expected_config_file.
    """
    report_file = artifacts_dir / "nvl_stock_report.json"
    if not report_file.exists():
        raise FileNotFoundError(f"Không tìm thấy nvl_stock_report.json trong thư mục artifacts: {artifacts_dir}")

    report_data = json.loads(report_file.read_text(encoding="utf-8"))
    actual_mode = report_data.get("mode")

    if expected_publish and actual_mode != "publish":
        raise ValueError(
            f"Publish mode không khớp: yêu cầu publish=True nhưng report ghi nhận mode='{actual_mode}'"
        )
    if not expected_publish and actual_mode == "publish":
        raise ValueError(
            f"Publish mode không khớp: yêu cầu dry-run (publish=False) nhưng report ghi nhận mode='publish' (nguy cơ ghi đè!)"
        )

    # Kiểm tra cấu hình đích
    cfg_path = Path(expected_config_file)
    if cfg_path.exists():
        expected_cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
        exp_tgt_path = expected_cfg.get("target", {}).get("sharepoint_path")
        exp_tgt_name = expected_cfg.get("target", {}).get("name")

        act_tgt_path = report_data.get("target", {}).get("sharepoint_path")
        act_tgt_name = report_data.get("target", {}).get("name")

        if exp_tgt_path and act_tgt_path and exp_tgt_path.strip().lower() != act_tgt_path.strip().lower():
            raise ValueError(
                f"Cấu hình đường dẫn đích không khớp: yêu cầu '{exp_tgt_path}' (từ {expected_config_file}) "
                f"nhưng report chạy với '{act_tgt_path}'"
            )
        if exp_tgt_name and act_tgt_name and exp_tgt_name.strip().lower() != act_tgt_name.strip().lower():
            raise ValueError(
                f"Cấu hình tên file đích không khớp: yêu cầu '{exp_tgt_name}' (từ {expected_config_file}) "
                f"nhưng report chạy với '{act_tgt_name}'"
            )

    audit_file = artifacts_dir / "audit_summary.json"
    if audit_file.exists():
        audit_data = json.loads(audit_file.read_text(encoding="utf-8"))
        if "config_file" in audit_data:
            act_cfg = Path(audit_data["config_file"]).name
            exp_cfg = Path(expected_config_file).name
            if act_cfg != exp_cfg:
                raise ValueError(
                    f"Config file trong audit_summary ('{act_cfg}') không khớp với '{exp_cfg}'"
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


def download_run_artifacts(token: str, repo: str, run_id: int, target_dir: Path) -> list[str]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    }
    artifacts_url = f"https://api.github.com/repos/{repo}/actions/runs/{run_id}/artifacts"
    resp = requests.get(artifacts_url, headers=headers, timeout=30)
    if resp.status_code != 200:
        raise RuntimeError(f"Lỗi lấy danh sách artifacts của run {run_id}: HTTP {resp.status_code}")

    artifacts = resp.json().get("artifacts", [])
    if not artifacts:
        print(f"[TRIGGER] Run {run_id} không có artifacts nào.")
        return []

    target_dir.mkdir(parents=True, exist_ok=True)
    downloaded_files = []

    for art in artifacts:
        art_id = art["id"]
        art_name = art["name"]
        size_bytes = art.get("size_in_bytes", 0)
        print(f"[TRIGGER] Đang tải artifact '{art_name}' ({size_bytes} bytes)...")

        down_url = f"https://api.github.com/repos/{repo}/actions/artifacts/{art_id}/zip"
        down_resp = requests.get(down_url, headers=headers, timeout=60)
        if down_resp.status_code != 200:
            print(f"[TRIGGER] Cảnh báo: Không thể tải artifact {art_name} (HTTP {down_resp.status_code})", file=sys.stderr)
            continue

        try:
            with zipfile.ZipFile(io.BytesIO(down_resp.content)) as z:
                z.extractall(target_dir)
                extracted = z.namelist()
                downloaded_files.extend(extracted)
                print(f"[TRIGGER] Đã giải nén: {extracted}")
        except zipfile.BadZipFile as bz:
            print(f"[TRIGGER] Cảnh báo: File tải về không phải zip hợp lệ: {bz}", file=sys.stderr)

    print(f"[TRIGGER] Toàn bộ artifacts đã lưu vào: {target_dir.resolve()}")
    return downloaded_files


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Kích hoạt và theo dõi dry-run workflow của NVL.")
    parser.add_argument("--sha", help="Commit SHA mong muốn kiểm tra (mặc định: git HEAD hiện tại).")
    parser.add_argument("--run-id", type=int, help="Theo dõi một run cụ thể thay vì tìm theo SHA.")
    parser.add_argument("--branch", default=BRANCH, help="Nhánh git.")
    parser.add_argument("--workflow", default=WORKFLOW, help="Tên file workflow.")
    parser.add_argument("--repo", default=REPO, help="Định dạng owner/repo trên GitHub.")
    parser.add_argument("--no-dispatch", action="store_true", help="Không dispatch mới, chỉ tìm run có sẵn của SHA.")
    parser.add_argument("--publish", action="store_true", help="Publish lên SharePoint.")
    parser.add_argument("--config-file", default="nvl_stock_config.json", help="Tên file cấu hình JSON.")
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
            seen_run_ids = get_existing_run_ids(token, args.repo, args.branch, args.workflow)
            print(f"[TRIGGER] Đã ghi nhận {len(seen_run_ids)} workflow runs hiện có trước dispatch.")
            try:
                inputs_payload = {
                    "publish": bool(args.publish),
                    "config_file": args.config_file,
                }
                dispatch_workflow(token, args.repo, args.workflow, args.branch, inputs_payload)
            except Exception as exc:
                print(f"[TRIGGER] LỖI DISPATCH: {exc}", file=sys.stderr)
                return 1

            # Sau dispatch: BẮT BUỘC tìm run MỚI (loại trừ seen_run_ids) và event == 'workflow_dispatch'
            try:
                run_data = find_run_for_sha(
                    token,
                    args.repo,
                    args.branch,
                    args.workflow,
                    expected_sha,
                    exclude_run_ids=seen_run_ids,
                    require_event="workflow_dispatch",
                    timeout_seconds=60,
                )
                run_id = run_data["id"]
                head_sha = run_data.get("head_sha", expected_sha)
            except Exception as exc:
                print(f"[TRIGGER] LỖI TÌM RUN: {exc}", file=sys.stderr)
                return 1
        else:
            # Chế độ no-dispatch: tìm run có sẵn
            try:
                run_data = find_run_for_sha(
                    token,
                    args.repo,
                    args.branch,
                    args.workflow,
                    expected_sha,
                    timeout_seconds=30,
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
            run_branch = run_data.get("head_branch", "")
            run_workflow = run_data.get("path", "") or run_data.get("name", "")

            if expected_sha and not (run_sha == expected_sha or run_sha.startswith(expected_sha)):
                print(
                    f"[TRIGGER] LỖI: Run {run_id} có SHA '{run_sha}' không khớp expected SHA '{expected_sha}'.",
                    file=sys.stderr,
                )
                return 1
            if run_branch != args.branch:
                print(
                    f"[TRIGGER] LỖI: Run {run_id} thuộc nhánh '{run_branch}' không khớp branch '{args.branch}'.",
                    file=sys.stderr,
                )
                return 1
            if not (args.workflow in run_workflow or run_data.get("name") == "Sync SharePoint NVL Stock"):
                print(
                    f"[TRIGGER] LỖI: Run {run_id} workflow '{run_workflow}' không khớp '{args.workflow}'.",
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

    # Xác minh artifact khớp với config_file và publish mode
    try:
        verify_run_execution(target_dir, args.config_file, args.publish)
        print(f"[TRIGGER] Đã xác minh thành công: Artifacts khớp đúng config '{args.config_file}' và mode (publish={args.publish}).")
    except Exception as ver_exc:
        print(f"[TRIGGER] LỖI XÁC MINH ARTIFACTS: {ver_exc}", file=sys.stderr)
        return 1

    print(f"SUCCESS: Workflow {run_id} hoàn tất thành công với conclusion: {conclusion}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
