#!/usr/bin/env python3
"""Script đảm bảo file bản sao staging Test_Ke_hoach_mua_hang_copy.xlsx tồn tại trên SharePoint.

Quy tắc an toàn:
1. Kiểm tra xem file bản sao đã tồn tại chưa bằng get_item_by_path.
2. Nếu ĐÃ TỒN TẠI: Tuyệt đối KHÔNG ghi đè, giữ nguyên nội dung hiện tại, chỉ trích xuất identity (item_id, eTag, sourcedoc).
3. Nếu CHƯA TỒN TẠI: Tải nội dung file chính thức (Kế hoạch mua hàng.xlsx) và tạo mới file bản sao qua create_file_by_path (conflict_behavior='fail').
4. Trích xuất chính xác sourcedoc (GUID) từ eTag của bản sao và ghi ra staging_copy_info.json và nvl_stock_report.json (mode='ensure_staging_copy').
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

# Add project root to sys.path
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from sync_stock import GraphClient, GraphRequestError, get_access_token


def parse_sourcedoc_from_etag(etag: str) -> str:
    """Trích xuất GUID sourcedoc từ eTag của SharePoint (định dạng '{GUID},version')."""
    if not etag:
        return ""
    m = re.search(r"\{([A-Fa-f0-9\-]+)\}", etag)
    return m.group(1).upper() if m else ""


def ensure_staging_copy(
    config_path: str = "nvl_stock_staging_config.json",
    official_config_path: str = "nvl_stock_config.json",
    out_dir_path: str = ".",
    graph: GraphClient | None = None,
) -> dict:
    out_dir = Path(out_dir_path)
    out_dir.mkdir(parents=True, exist_ok=True)

    # 1. Đọc config staging và official
    cfg_file = Path(config_path)
    if not cfg_file.exists():
        raise FileNotFoundError(f"File config staging không tồn tại: {config_path}")
    with open(cfg_file, "r", encoding="utf-8") as f:
        staging_cfg = json.load(f)

    official_file = Path(official_config_path)
    if not official_file.exists():
        raise FileNotFoundError(f"File config chính thức không tồn tại: {official_config_path}")
    with open(official_file, "r", encoding="utf-8") as f:
        official_cfg = json.load(f)

    staging_target = staging_cfg.get("target", {})
    staging_path = staging_target.get("sharepoint_path", "Tinh san xuat Mua hang 2027/Test_Ke_hoach_mua_hang_copy.xlsx")
    staging_name = staging_target.get("name", "Test_Ke_hoach_mua_hang_copy.xlsx")

    official_target = official_cfg.get("target", {})
    official_path = official_target.get("sharepoint_path", "Tinh san xuat Mua hang 2027/Kế hoạch mua hàng.xlsx")

    # 2. Khởi tạo Graph client
    if graph is None:
        token = get_access_token()
        graph = GraphClient(token)

    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    print(f"[ENSURE_COPY] Kiểm tra file bản sao '{staging_path}' trên SharePoint...")
    target_item = None
    status = "unknown"

    try:
        target_item = graph.get_item_by_path(drive_id, staging_path)
        status = "existed"
        print(f"[ENSURE_COPY] File bản sao ĐÃ TỒN TẠI (ID: {target_item.get('id')}). Không ghi đè.")
    except GraphRequestError as exc:
        if exc.status_code == 404:
            print(f"[ENSURE_COPY] File bản sao chưa tồn tại. Bắt đầu tạo mới từ file gốc: '{official_path}'...")
            official_item = graph.get_item_by_path(drive_id, official_path)
            official_bytes = graph.download_file(drive_id, official_item["id"])
            print(f"[ENSURE_COPY] Đã tải file gốc ({len(official_bytes)} bytes). Đang tạo '{staging_path}'...")

            graph.create_file_by_path(
                drive_id,
                staging_path,
                official_bytes,
                conflict_behavior="fail",
            )
            # Re-fetch item to ensure complete item metadata
            target_item = graph.get_item_by_path(drive_id, staging_path)
            status = "created"
            print(f"[ENSURE_COPY] Đã tạo thành công file bản sao (ID: {target_item.get('id')}).")
        else:
            raise

    target_etag = target_item.get("eTag", "")
    sourcedoc = parse_sourcedoc_from_etag(target_etag)
    item_id = target_item.get("id", "")
    name = target_item.get("name", staging_name)
    size = target_item.get("size", 0)
    web_url = target_item.get("webUrl", "")

    info = {
        "status": status,
        "name": name,
        "sharepoint_path": staging_path,
        "item_id": item_id,
        "eTag": target_etag,
        "sourcedoc": sourcedoc,
        "size": size,
        "web_url": web_url,
        "timestamp": datetime.now(timezone.utc).isoformat(),
    }

    # Ghi file staging_copy_info.json
    info_path = out_dir / "staging_copy_info.json"
    with open(info_path, "w", encoding="utf-8") as f:
        json.dump(info, f, indent=2, ensure_ascii=False)
    print(f"[ENSURE_COPY] Đã lưu thông tin bản sao vào: {info_path}")

    # Ghi nvl_stock_report.json với mode 'ensure_staging_copy' để hỗ trợ verify artifact thống nhất
    report = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "status": "success",
        "mode": "ensure_staging_copy",
        "staging_copy": info,
        "source": staging_cfg.get("source", {}),
        "target": {
            "name": name,
            "sharepoint_path": staging_path,
            "sourcedoc": sourcedoc,
            "item_id": item_id,
            "eTag": target_etag,
        },
    }
    report_path = out_dir / "nvl_stock_report.json"
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)

    print("=" * 60)
    print("THÔNG TIN BẢN SAO STAGING SHAREPOINT:")
    print(f"  - Trạng thái:      {status.upper()}")
    print(f"  - Tên file:        {name}")
    print(f"  - Đường dẫn:       {staging_path}")
    print(f"  - Item ID:         {item_id}")
    print(f"  - eTag:            {target_etag}")
    print(f"  - Sourcedoc GUID:  {sourcedoc}")
    print("=" * 60)

    return info


def main():
    parser = argparse.ArgumentParser(
        description="Đảm bảo file bản sao staging Test_Ke_hoach_mua_hang_copy.xlsx tồn tại trên SharePoint."
    )
    parser.add_argument("--config", default="nvl_stock_staging_config.json", help="Đường dẫn config staging")
    parser.add_argument("--official-config", default="nvl_stock_config.json", help="Đường dẫn config chính thức")
    parser.add_argument("--out-dir", default=".", help="Thư mục ghi kết quả")
    args = parser.parse_args()

    try:
        ensure_staging_copy(
            config_path=args.config,
            official_config_path=args.official_config,
            out_dir_path=args.out_dir,
        )
        return 0
    except Exception as exc:
        print(f"[ENSURE_COPY] LỖI: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
