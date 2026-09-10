"""Trigger dry-run on sync-nvl-stock.yml, monitor status and conclusion,
download artifacts for diagnostics, and fail with exit code 1 if conclusion != success.

Strictly matches expected commit SHA or run ID to prevent using stale artifacts as evidence.
"""

import argparse
import hashlib
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
    """Xác minh artifact sau khi tải từ run:
    1. Yêu cầu config kỳ vọng phải tồn tại.
    2. Yêu cầu mode trong report hợp lệ và khớp với expected_publish.
    3. Yêu cầu status trong report không phải 'failed'.
    4. Yêu cầu đủ thông tin nguồn và đích trong report (name, sharepoint_path) và phải khớp cấu hình.
    5. Nếu có audit_summary.json, kiểm tra status và config_file.
    """
    # 1. Kiểm tra config tồn tại
    cfg_path = Path(expected_config_file)
    if not cfg_path.exists():
        raise FileNotFoundError(f"File cấu hình kỳ vọng '{expected_config_file}' không tồn tại.")
    expected_cfg = json.loads(cfg_path.read_text(encoding="utf-8"))

    # 2. Kiểm tra report tồn tại trong thư mục artifacts vừa tải
    report_file = artifacts_dir / "nvl_stock_report.json"
    if not report_file.exists():
        raise FileNotFoundError(f"Không tìm thấy nvl_stock_report.json trong thư mục artifacts: {artifacts_dir}")

    report_data = json.loads(report_file.read_text(encoding="utf-8"))

    # 3. Kiểm tra mode hợp lệ
    actual_mode = report_data.get("mode")
    valid_modes = {"dry_run", "publish", "offline", "failed", "ensure_staging_copy"}
    if actual_mode not in valid_modes:
        raise ValueError(f"Report có mode không hợp lệ: {actual_mode!r}. Mode hợp lệ: {valid_modes}")

    if expected_publish == "ensure_staging_copy":
        if actual_mode != "ensure_staging_copy":
            raise ValueError(
                f"Mode không khớp: yêu cầu mode='ensure_staging_copy' nhưng report ghi nhận mode='{actual_mode}'"
            )
    elif expected_publish and actual_mode != "publish":
        raise ValueError(
            f"Publish mode không khớp: yêu cầu publish=True nhưng report ghi nhận mode='{actual_mode}'"
        )
    elif not expected_publish and actual_mode == "publish":
        raise ValueError(
            f"Publish mode không khớp: yêu cầu dry-run (publish=False) nhưng report ghi nhận mode='publish' (nguy cơ ghi đè!)"
        )

    # 4. Kiểm tra status report
    if report_data.get("status") == "failed":
        raise ValueError(
            f"Report ghi nhận trạng thái thất bại: phase='{report_data.get('phase')}', "
            f"error='{report_data.get('error_message')}'"
        )

    # 5. Yêu cầu đủ thông tin nguồn và đích để đối chiếu
    source_info = report_data.get("source")
    target_info = report_data.get("target")
    if not isinstance(source_info, dict) or not isinstance(target_info, dict):
        raise ValueError("Report thiếu thông tin bắt buộc về 'source' hoặc 'target' (phải là đối tượng dict).")

    act_src_name = source_info.get("name")
    act_src_path = source_info.get("sharepoint_path")
    act_tgt_name = target_info.get("name")
    act_tgt_path = target_info.get("sharepoint_path")

    if not act_src_name or not act_src_path:
        raise ValueError("Report thiếu trường bắt buộc của nguồn: 'name' hoặc 'sharepoint_path'.")
    if not act_tgt_name or not act_tgt_path:
        raise ValueError("Report thiếu trường bắt buộc của đích: 'name' hoặc 'sharepoint_path'.")

    exp_src = expected_cfg.get("source", {})
    exp_tgt = expected_cfg.get("target", {})
    exp_src_path = exp_src.get("sharepoint_path")
    exp_src_name = exp_src.get("name")
    exp_tgt_path = exp_tgt.get("sharepoint_path")
    exp_tgt_name = exp_tgt.get("name")

    if exp_src_path and act_src_path and exp_src_path.strip().lower() != act_src_path.strip().lower():
        raise ValueError(
            f"Cấu hình đường dẫn nguồn không khớp: yêu cầu '{exp_src_path}' (từ {expected_config_file}) "
            f"nhưng report chạy với '{act_src_path}'"
        )
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

    # 6. Kiểm tra audit_summary.json nếu có
    audit_file = artifacts_dir / "audit_summary.json"
    if audit_file.exists():
        audit_data = json.loads(audit_file.read_text(encoding="utf-8"))
        if audit_data.get("status") == "failed":
            raise ValueError(
                f"Audit summary ghi nhận thất bại: phase='{audit_data.get('phase')}', "
                f"error='{audit_data.get('error_message')}'"
            )
        if "config_file" in audit_data:
            act_cfg = Path(audit_data["config_file"]).name
            exp_cfg = Path(expected_config_file).name
            if act_cfg != exp_cfg:
                raise ValueError(
                    f"Config file trong audit_summary ('{act_cfg}') không khớp với '{exp_cfg}'"
                )

    # 7. Kiểm tra official_target_backup_info.json và backup raw nếu có
    info_file = artifacts_dir / "official_target_backup_info.json"
    raw_file = artifacts_dir / "official_backup_target_raw.xlsx"
    if info_file.exists():
        info_data = json.loads(info_file.read_text(encoding="utf-8"))
        exp_sha = info_data.get("sha256")
        if not raw_file.exists():
            raise FileNotFoundError(
                "Artifact chứa 'official_target_backup_info.json' nhưng thiếu file backup raw 'official_backup_target_raw.xlsx'."
            )
        act_sha = hashlib.sha256(raw_file.read_bytes()).hexdigest()
        if exp_sha and act_sha.lower() != exp_sha.lower():
            raise ValueError(
                f"SHA-256 của official_backup_target_raw.xlsx ({act_sha}) không khớp với official_target_backup_info.json ({exp_sha})!"
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
    base_target_dir: Path,
    *,
    require_report: bool = True,
) -> tuple[Path, list[str]]:
    """Tải artifacts vào thư mục riêng biệt theo run ID:
    - Nếu không có artifact nào: raise RuntimeError.
    - Nếu tải lỗi hoặc ZIP hỏng: raise RuntimeError.
    - Nếu require_report=True và thiếu nvl_stock_report.json: raise FileNotFoundError.
    - Trả về tuple (run_artifacts_dir, downloaded_files).
    """
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
        raise RuntimeError(f"Run {run_id} không có artifact nào trên GitHub.")

    # Tạo thư mục riêng biệt theo run_id và dọn dẹp các file cũ trong thư mục này
    run_dir = base_target_dir / f"run_{run_id}"
    if run_dir.exists():
        for p in run_dir.iterdir():
            if p.is_file():
                try:
                    p.unlink()
                except OSError:
                    pass
    else:
        run_dir.mkdir(parents=True, exist_ok=True)

    downloaded_files = []

    for art in artifacts:
        art_id = art.get("id")
        art_name = art.get("name", f"artifact_{art_id}")
        size_bytes = art.get("size_in_bytes", 0)
        print(f"[TRIGGER] Đang tải artifact '{art_name}' ({size_bytes} bytes)...")

        down_url = f"https://api.github.com/repos/{repo}/actions/artifacts/{art_id}/zip"
        down_resp = requests.get(down_url, headers=headers, timeout=60)
        if down_resp.status_code != 200:
            raise RuntimeError(
                f"Tải artifact '{art_name}' của run {run_id} thất bại (HTTP {down_resp.status_code})"
            )

        try:
            with zipfile.ZipFile(io.BytesIO(down_resp.content)) as z:
                z.extractall(run_dir)
                extracted = z.namelist()
                downloaded_files.extend(extracted)
                print(f"[TRIGGER] Đã giải nén vào {run_dir}: {extracted}")
        except zipfile.BadZipFile as bz:
            raise RuntimeError(f"Artifact '{art_name}' của run {run_id} không phải file zip hợp lệ: {bz}")

    if require_report:
        report_file = run_dir / "nvl_stock_report.json"
        if not report_file.exists():
            raise FileNotFoundError(
                f"Artifact vừa tải của run {run_id} không chứa file bắt buộc 'nvl_stock_report.json'."
            )

    print(f"[TRIGGER] Toàn bộ artifacts của run {run_id} đã lưu vào thư mục riêng: {run_dir.resolve()}")
    return run_dir, downloaded_files


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Kích hoạt và theo dõi dry-run workflow của NVL.")
    parser.add_argument("--sha", help="Commit SHA mong muốn kiểm tra (mặc định: git HEAD hiện tại).")
    parser.add_argument("--run-id", type=int, help="Theo dõi một run cụ thể thay vì tìm theo SHA.")
    parser.add_argument("--branch", default=BRANCH, help="Nhánh git.")
    parser.add_argument("--workflow", default=WORKFLOW, help="Tên file workflow.")
    parser.add_argument("--repo", default=REPO, help="Định dạng owner/repo trên GitHub.")
    parser.add_argument("--no-dispatch", action="store_true", help="Không dispatch mới, chỉ tìm run có sẵn của SHA.")
    parser.add_argument("--publish", action="store_true", help="Publish lên SharePoint.")
    parser.add_argument("--ensure-staging-copy", action="store_true", help="Kiểm tra hoặc tạo file bản sao staging Test_Ke_hoach_mua_hang_copy.xlsx trên SharePoint.")
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
                    "ensure_staging_copy": bool(args.ensure_staging_copy),
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

    # Tải artifacts vào thư mục riêng theo run ID
    target_dir = Path(args.out_dir)
    try:
        run_artifacts_dir, downloaded_files = download_run_artifacts(token, args.repo, run_id, target_dir)
    except Exception as down_exc:
        print(f"[TRIGGER] LỖI TẢI ARTIFACTS: {down_exc}", file=sys.stderr)
        return 1

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
    expected_verify_mode = "ensure_staging_copy" if args.ensure_staging_copy else args.publish
    try:
        verify_run_execution(run_artifacts_dir, args.config_file, expected_verify_mode)
        print(f"[TRIGGER] Đã xác minh thành công: Artifacts khớp đúng config '{args.config_file}' và mode ({expected_verify_mode}).")
    except Exception as ver_exc:
        print(f"[TRIGGER] LỖI XÁC MINH ARTIFACTS: {ver_exc}", file=sys.stderr)
        return 1

    info_file = run_artifacts_dir / "staging_copy_info.json"
    if info_file.exists():
        try:
            copy_info = json.loads(info_file.read_text(encoding="utf-8"))
            print("\n" + "=" * 60)
            print("THÔNG TIN BẢN SAO STAGING SHAREPOINT VỪA TẢI VỀ:")
            print(f"  - Trạng thái:      {copy_info.get('status', '').upper()}")
            print(f"  - Tên file:        {copy_info.get('name')}")
            print(f"  - Đường dẫn:       {copy_info.get('sharepoint_path')}")
            print(f"  - Item ID:         {copy_info.get('item_id')}")
            print(f"  - eTag:            {copy_info.get('eTag')}")
            print(f"  - Sourcedoc GUID:  {copy_info.get('sourcedoc')}")
            print("=" * 60 + "\n")
        except Exception:
            pass

    print(f"SUCCESS: Workflow {run_id} hoàn tất thành công với conclusion: {conclusion}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
