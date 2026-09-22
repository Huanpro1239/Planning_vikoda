"""Persistent input fingerprint state for finished-goods planning."""

import json
from pathlib import Path


STATE_FILE = Path("state.json")
SYNC_VERSION = 3


def load_state(*, path=None):
    state_file = Path(path) if path is not None else STATE_FILE
    if not state_file.exists():
        return {}

    try:
        data = json.loads(state_file.read_text(encoding="utf-8"))
    except Exception:
        return {}

    if "sources" not in data and data.get("source_etag"):
        return {
            "sources": {
                "actual_stock": data["source_etag"],
            }
        }

    return data


def save_state(
    source_etags,
    conversion_hash,
    fc_hash=None,
    no_kho_hash=None,
    planning_inputs_hash=None,
    engine_version=None,
    *,
    path=None,
    **kwargs,
):
    payload = {
        "sync_version": SYNC_VERSION,
        "conversion_hash": conversion_hash,
        "sources": source_etags,
    }
    if fc_hash is not None:
        payload["fc_hash"] = fc_hash
    if no_kho_hash is not None:
        payload["no_kho_hash"] = no_kho_hash
    if planning_inputs_hash is not None:
        payload["planning_inputs_hash"] = planning_inputs_hash
    if engine_version is not None:
        payload["engine_version"] = engine_version
    for key, value in kwargs.items():
        if value is not None:
            payload[key] = value

    state_file = Path(path) if path is not None else STATE_FILE
    state_file.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
