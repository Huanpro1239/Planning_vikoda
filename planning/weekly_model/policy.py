"""Policy metadata normalization for the weekly planning adapter."""

from __future__ import annotations

import re
from typing import Any


DEBT_HEADER_NAMES: frozenset[str] = frozenset({
    "debt mode",
    "cách tính nợ",
    "cach tinh no",
})
PROFILE_HEADER_NAMES: frozenset[str] = frozenset({
    "schedule profile",
    "profile lịch",
    "profile lich",
    "lịch sản xuất",
    "lich san xuat",
})


def norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def find_header_col(
    headers: list[Any],
    names: set[str] | frozenset[str],
) -> int | None:
    """Return a 0-indexed matching header column."""
    wanted = {norm(name) for name in names}
    for idx, header in enumerate(headers):
        if norm(header) in wanted:
            return idx
    return None


def header_col(
    worksheet,
    names: set[str] | frozenset[str],
) -> int | None:
    headers = [
        worksheet.cell(1, col).value
        for col in range(1, worksheet.max_column + 1)
    ]
    idx = find_header_col(headers, names)
    return idx + 1 if idx is not None else None


def debt_mode(value: Any) -> str | None:
    text = norm(value).replace("-", "_").replace(" ", "_")
    if not text:
        return None
    if text in {
        "subtract_book_on_debt",
        "tru_ton_so",
        "trừ_tồn_sổ",
    }:
        return "SUBTRACT_BOOK_ON_DEBT"
    if text in {
        "ignore_book_on_debt",
        "khong_tru_ton_so",
        "không_trừ_tồn_sổ",
    }:
        return "IGNORE_BOOK_ON_DEBT"
    raise RuntimeError(f"Debt mode không hợp lệ: {value!r}")


def normalize_debt_mode(value: Any) -> str | None:
    """Normalize debt mode safely for engine and fingerprint usage."""
    try:
        return debt_mode(value)
    except Exception:
        return str(value or "").strip() or None


def schedule_profile(value: Any) -> str | None:
    text = norm(value).replace("-", "_").replace(" ", "_")
    if not text:
        return None
    if text in {
        "continuous",
        "lien_tuc",
        "liên_tục",
    }:
        return "CONTINUOUS"
    if text in {
        "spread_non_sunday",
        "rai_khong_chu_nhat",
        "rải_không_chủ_nhật",
    }:
        return "SPREAD_NON_SUNDAY"
    raise RuntimeError(
        f"Schedule profile không hợp lệ: {value!r}"
    )


def normalize_profile(value: Any) -> str | None:
    """Normalize schedule profile safely for engine and fingerprint usage."""
    try:
        return schedule_profile(value)
    except Exception:
        return str(value or "").strip() or None
