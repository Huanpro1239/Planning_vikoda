"""Deterministic openpyxl resource cleanup.

Keeping the input BytesIO alive until every workbook archive is closed avoids
Python 3.12 ZipFile.__del__ warnings such as "I/O operation on closed file".
"""

from __future__ import annotations

from contextlib import contextmanager
from io import BytesIO
from typing import Iterator

from openpyxl import load_workbook


def _close_zip_archive(archive) -> None:
    if archive is None:
        return
    try:
        archive.close()
    except Exception:
        pass
    try:
        archive.fp = None
    except Exception:
        pass


def safe_close_workbook(workbook) -> None:
    if workbook is None:
        return

    archives = []
    for attr in ("_archive", "vba_archive"):
        try:
            archive = getattr(workbook, attr, None)
        except Exception:
            archive = None
        if archive is not None and all(archive is not item for item in archives):
            archives.append(archive)

    try:
        workbook.close()
    except Exception:
        pass

    for archive in archives:
        _close_zip_archive(archive)


@contextmanager
def open_workbook_bytes(data: bytes, **kwargs) -> Iterator:
    buffer = BytesIO(data)
    workbook = None
    try:
        workbook = load_workbook(buffer, **kwargs)
        yield workbook
    finally:
        safe_close_workbook(workbook)
        try:
            buffer.close()
        except Exception:
            pass


__all__ = ["safe_close_workbook", "open_workbook_bytes"]
