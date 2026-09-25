"""Canonical API and package-native CLI for NVL stock synchronization."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from excel.openpyxl_io import safe_close_workbook as _safe_close_workbook
from nvl.config import DEFAULT_CONFIG_FILE, load_nvl_config
from nvl.models import NVLConfig, NVLReconcileResult
from nvl.reconcile import (
    check_target_sheet_safety,
    read_nvl_source_stock,
    reconcile_nvl_target,
)
from nvl.reporting import generate_nvl_error_report, generate_nvl_report
from nvl.service import run_nvl_sync
from nvl.values import normalize_nvl_code, parse_nvl_quantity
from nvl.workbook import (
    identify_sharepoint_metadata_exemption,
    patch_nvl_destination_workbook,
    verify_nvl_patched_workbook,
)

def main():
    parser = argparse.ArgumentParser(
        description="Đồng bộ tồn kho nguyên vật liệu (NVL) từ XNT_ketoan_Vikoda sang Kế hoạch mua hàng."
    )
    parser.add_argument("--config", default=DEFAULT_CONFIG_FILE, help="Đường dẫn file cấu hình JSON.")
    parser.add_argument("--source-file", help="File nguồn cục bộ (chế độ offline).")
    parser.add_argument("--target-file", help="File đích cục bộ (chế độ offline).")
    parser.add_argument("--target-path", help="Ghi đè đường dẫn file đích trên SharePoint.")
    parser.add_argument("--out", help="Thư mục xuất proposal và report.")
    parser.add_argument("--publish", action="store_true", help="Publish lên SharePoint (mặc định là dry-run).")

    args = parser.parse_args()

    out_dir = Path(args.out) if args.out else (Path("offline_out/nvl") if (args.source_file or args.target_file) else Path("."))

    sync_started = False
    try:
        cfg = load_nvl_config(args.config)
        if args.target_path:
            cfg.target_path = args.target_path

        sync_started = True
        rep = run_nvl_sync(
            cfg,
            source_file=args.source_file,
            target_file=args.target_file,
            out_dir=args.out,
            publish=args.publish,
        )
        print(f"Hoàn thành ({rep.get('status')}): {rep.get('message')}")
    except Exception as exc:
        report_file = out_dir / "nvl_stock_report.json"
        proposal_file = out_dir / "nvl_stock_proposal.xlsx"

        if not sync_started:
            # Lỗi xảy ra TRƯỚC KHI sync khởi động (ví dụ load config lỗi, thiếu file config, JSON hỏng):
            # 1. Xóa proposal cũ từ lần chạy trước nếu có để tránh proposal không hợp lệ tồn tại
            try:
                if proposal_file.exists():
                    proposal_file.unlink(missing_ok=True)
            except OSError:
                pass

            # 2. Luôn ghi đè báo cáo lỗi hiện tại (phase="init_cli", status="failed") thay thế report cũ
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
                err_rep = generate_nvl_error_report(
                    None,
                    mode="publish" if args.publish else ("offline" if (args.source_file or args.target_file) else "dry_run"),
                    phase="init_cli",
                    attempt=1,
                    error=exc,
                )
                report_file.write_text(json.dumps(err_rep, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            except OSError:
                pass
        else:
            # Nếu sync đã khởi động, run_nvl_sync đã chủ động ghi báo cáo lỗi chi tiết chuyên sâu (như post_upload_verify)
            # Chỉ ghi fallback nếu report_file vì lý do nào đó chưa tồn tại
            if not report_file.exists():
                try:
                    out_dir.mkdir(parents=True, exist_ok=True)
                    err_rep = generate_nvl_error_report(
                        None,
                        mode="publish" if args.publish else ("offline" if (args.source_file or args.target_file) else "dry_run"),
                        phase="cli_unhandled",
                        attempt=1,
                        error=exc,
                    )
                    report_file.write_text(json.dumps(err_rep, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
                except OSError:
                    pass

        print(f"LỖI: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()


__all__ = [
    "NVLConfig",
    "NVLReconcileResult",
    "load_nvl_config",
    "normalize_nvl_code",
    "parse_nvl_quantity",
    "check_target_sheet_safety",
    "read_nvl_source_stock",
    "reconcile_nvl_target",
    "patch_nvl_destination_workbook",
    "identify_sharepoint_metadata_exemption",
    "verify_nvl_patched_workbook",
    "generate_nvl_report",
    "generate_nvl_error_report",
    "run_nvl_sync",
]
