"""Runtime-state helpers for Planning metrics."""

import calendar
import json
from pathlib import Path


RUNTIME_STATE_FILE = Path("planning_runtime.json")
RUNTIME_STATE_VERSION = 1


def load_runtime_state():
    if not RUNTIME_STATE_FILE.exists():
        return {
            "version": RUNTIME_STATE_VERSION,
            "opening_debt_by_month": {},
            "opening_consignment_by_month": {},
        }

    try:
        state = json.loads(RUNTIME_STATE_FILE.read_text(encoding="utf-8"))
    except Exception as exc:
        raise RuntimeError(
            f"Không đọc được {RUNTIME_STATE_FILE}: {exc}"
        ) from exc

    state.setdefault("version", RUNTIME_STATE_VERSION)
    state.setdefault("opening_debt_by_month", {})
    state.setdefault("opening_consignment_by_month", {})
    return state


def save_runtime_state(state):
    state["version"] = RUNTIME_STATE_VERSION
    RUNTIME_STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def month_key(year, month):
    return f"{year:04d}-{month:02d}"


def next_month(year, month):
    if month == 12:
        return year + 1, 1
    return year, month + 1


def is_month_end(value):
    return value.day == calendar.monthrange(value.year, value.month)[1]
