"""SharePoint snapshot acquisition and provenance."""

import hashlib

import stock

from .constants import SOURCES


def read_snapshot(graph, drive_id):
    target_item = graph.get_item_by_path(drive_id, stock.DEST_PATH)
    target_bytes = graph.download_file(drive_id, target_item["id"])

    source_bytes = {}
    source_items = {}
    for key, path in SOURCES.items():
        item = graph.get_item_by_path(drive_id, path)
        data = graph.download_file(drive_id, item["id"])
        source_items[key] = item
        source_bytes[key] = data

    revision = {
        "target": {
            "path": stock.DEST_PATH,
            "item_id": target_item.get("id"),
            "etag": target_item.get("eTag"),
            "last_modified": target_item.get("lastModifiedDateTime"),
            "sha256": hashlib.sha256(target_bytes).hexdigest(),
        },
        "sources": {
            key: {
                "path": SOURCES[key],
                "item_id": source_items[key].get("id"),
                "etag": source_items[key].get("eTag"),
                "last_modified": source_items[key].get("lastModifiedDateTime"),
                "sha256": hashlib.sha256(source_bytes[key]).hexdigest(),
            }
            for key in SOURCES
        },
    }
    return target_item, target_bytes, source_items, source_bytes, revision
