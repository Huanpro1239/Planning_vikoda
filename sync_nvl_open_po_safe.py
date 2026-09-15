"""Safe entrypoint cho đồng bộ Ton_NVL!E.

Resolver đặc thù NVL chỉ còn nhiệm vụ đọc cấu hình và gắn canonical SharePoint
helpers vào module nghiệp vụ. Encode URL, retry Graph GET và identity validation
được dùng chung từ ``sharepoint.client``.
"""

from __future__ import annotations

import json
from pathlib import Path
import sys

import sync_nvl_open_po as base
from sharepoint.client import (
    get_item_metadata,
    graph_share_id as _graph_share_id,
    resolve_share_url,
)


_ORIGINAL_RESOLVE_SOURCE_ITEM = base.resolve_source_item
READ_ATTEMPTS = 3


def _source_config_path_from_argv() -> Path:
    """Đọc đúng --source-config mà base.main() sẽ sử dụng."""
    args = sys.argv[1:]
    for index, arg in enumerate(args):
        if arg == "--source-config" and index + 1 < len(args):
            return Path(args[index + 1])
        if arg.startswith("--source-config="):
            return Path(arg.split("=", 1)[1])
    return Path(base.DEFAULT_SOURCE_CONFIG)


def _load_source_config() -> dict:
    path = _source_config_path_from_argv()
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("source", {})


def _configured_source_url() -> str:
    return str(_load_source_config().get("source_url", "")).strip()


def _item_metadata_with_retry(graph, drive_id: str, item_id: str) -> dict:
    return get_item_metadata(
        graph,
        drive_id,
        item_id,
        attempts=READ_ATTEMPTS,
    )


def resolve_source_item_exact_url(graph, drive_id: str, config) -> dict:
    """Resolve source bằng exact SharePoint URL; không Search theo tên, không pin item ID."""
    source_url = _configured_source_url()
    if not source_url:
        # Tương thích config cũ; production config hiện dùng source_url exact.
        return _ORIGINAL_RESOLVE_SOURCE_ITEM(graph, drive_id, config)

    item = resolve_share_url(
        graph,
        source_url,
        expected_name=config.source_name,
        expected_sourcedoc=config.source_sourcedoc,
        expected_drive_id=drive_id,
        attempts=READ_ATTEMPTS,
        label="source_url",
    )

    print(
        f"[Ton_NVL!E] Đã resolve đúng nguồn {config.source_name} từ exact SharePoint URL; "
        f"item_id={item['id']}; sourcedoc đã xác minh."
    )
    return item


def install_safe_resolver() -> None:
    # Mọi metadata GET sau resolve (sau download / trước upload) dùng retry chung.
    base._item_metadata = _item_metadata_with_retry
    base.resolve_source_item = resolve_source_item_exact_url


def main() -> None:
    install_safe_resolver()
    base.main()


if __name__ == "__main__":
    main()
