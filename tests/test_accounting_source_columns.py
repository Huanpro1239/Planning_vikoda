"""Regression tests for distinct VKD/Vikoda accounting source layouts.

The VKD 2026-10 export has ending stock in K, converted quantity in L,
and warehouse name in M. Vikoda's export instead has ending stock in M.
"""
import ast
from io import BytesIO
from pathlib import Path
import unittest

from openpyxl import Workbook

from stock.readers import read_single_value_source


ROOT = Path(__file__).resolve().parents[1]
CALLERS = ("planning/pipeline.py", "stock/service.py")


def _source_kwargs(path, source_key):
    tree = ast.parse((ROOT / path).read_text(encoding="utf-8"))
    matching_calls = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call) or not node.args:
            continue
        fn_name = (
            node.func.attr if isinstance(node.func, ast.Attribute)
            else node.func.id if isinstance(node.func, ast.Name)
            else None
        )
        if fn_name != "read_single_value_source":
            continue
        source = node.args[0]
        if not (
            isinstance(source, ast.Subscript)
            and isinstance(source.value, ast.Name)
            and source.value.id == "source_bytes"
            and isinstance(source.slice, ast.Constant)
            and source.slice.value == source_key
        ):
            continue
        matching_calls.append(node)

    if len(matching_calls) != 1:
        raise AssertionError(
            f"{path} must contain exactly one {source_key} reader call "
            f"(found {len(matching_calls)})"
        )
    return {
        kw.arg: ast.literal_eval(kw.value)
        for kw in matching_calls[0].keywords
    }


def _source_workbook(*, end_stock_column, other_stock_column, other_value):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"
    worksheet["B9"] = "Mã vật tư"
    worksheet["B11"] = "230100008"
    worksheet[f"{end_stock_column}11"] = 46738
    worksheet[f"{other_stock_column}11"] = other_value
    worksheet["B12"] = "230100014"
    worksheet[f"{end_stock_column}12"] = 0
    worksheet[f"{other_stock_column}12"] = other_value
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class AccountingSourceColumnTests(unittest.TestCase):
    def test_vkd_reads_ending_stock_k_not_converted_l_or_warehouse_m(self):
        source = _source_workbook(
            end_stock_column="K",
            other_stock_column="M",
            other_value="VKD - Kho Hàng Hóa",
        )
        # L is *already* converted; applying the target factor to L
        # again would double-convert the quantity.
        for path in CALLERS:
            with self.subTest(path=path):
                kwargs = _source_kwargs(path, "accounting_vkd")
                self.assertEqual(kwargs["value_column"], 11)
                self.assertEqual(kwargs["value_column_letter"], "K")
                self.assertTrue(kwargs["vkd_to_vikoda"])
                self.assertEqual(kwargs["source_name"], "XNT_ketoan_VKD.xlsm")
                self.assertEqual(
                    read_single_value_source(source, **kwargs),
                    {"130100008": 46738, "130100014": 0},
                )

    def test_vikoda_continues_to_read_ending_stock_m(self):
        source = _source_workbook(
            end_stock_column="M",
            other_stock_column="K",
            other_value=123,
        )
        for path in CALLERS:
            with self.subTest(path=path):
                kwargs = _source_kwargs(path, "accounting_vikoda")
                self.assertEqual(kwargs["value_column"], 13)
                self.assertEqual(kwargs["value_column_letter"], "M")
                self.assertFalse(kwargs.get("vkd_to_vikoda", False))
                self.assertEqual(
                    read_single_value_source(source, **kwargs),
                    {"230100008": 46738, "230100014": 0},
                )

    def test_vkd_invalid_stock_value_is_not_silently_zeroed(self):
        source = _source_workbook(
            end_stock_column="K",
            other_stock_column="M",
            other_value="VKD - Kho Hàng Hóa",
        )
        # Control: M is text, and an accidental regression to M must fail.
        kwargs = _source_kwargs(CALLERS[0], "accounting_vkd")
        kwargs.update(value_column=13, value_column_letter="M")
        with self.assertRaisesRegex(ValueError, "M11"):
            read_single_value_source(source, **kwargs)


if __name__ == "__main__":
    unittest.main()
