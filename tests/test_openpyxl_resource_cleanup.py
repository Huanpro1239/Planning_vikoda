"""Regression tests for deterministic openpyxl/ZipFile cleanup."""

import gc
import io
from contextlib import redirect_stderr
import unittest

from openpyxl import Workbook

from excel.openpyxl_io import open_workbook_bytes


def workbook_bytes() -> bytes:
    wb = Workbook()
    wb.active["A1"] = "ok"
    out = io.BytesIO()
    wb.save(out)
    wb.close()
    return out.getvalue()


class OpenpyxlResourceCleanupTests(unittest.TestCase):
    def test_read_only_keep_vba_cleanup_emits_no_zipfile_del_warning(self):
        captured = io.StringIO()
        with redirect_stderr(captured):
            with open_workbook_bytes(
                workbook_bytes(),
                data_only=True,
                read_only=True,
                keep_vba=True,
            ) as wb:
                self.assertEqual(wb.active["A1"].value, "ok")
            gc.collect()

        self.assertNotIn("ZipFile.__del__", captured.getvalue())
        self.assertNotIn("I/O operation on closed file", captured.getvalue())


if __name__ == "__main__":
    unittest.main()
