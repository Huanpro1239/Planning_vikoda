"""Unit tests for sync_nvl_stock module."""

from io import BytesIO
import json
from pathlib import Path
import tempfile
from typing import Any
import unittest
import zipfile
from lxml import etree
from openpyxl import Workbook, load_workbook

from sync_nvl_stock import (
    NVLConfig,
    normalize_nvl_code,
    parse_nvl_quantity,
    patch_nvl_destination_workbook,
    read_nvl_source_stock,
    reconcile_nvl_target,
    verify_nvl_patched_workbook,
)


def make_mock_config() -> NVLConfig:
    return NVLConfig(
        source_name="XNT_ketoan_Vikoda.xlsm",
        source_path="path/to/XNT_ketoan_Vikoda.xlsm",
        source_sheet="Sheet1",
        source_code_col=2,
        source_code_col_letter="B",
        source_value_col=13,
        source_value_col_letter="M",
        source_start_row=2,
        source_sourcedoc="SOURCE-GUID",
        target_name="Kế hoạch mua hàng.xlsx",
        target_path="path/to/Kế hoạch mua hàng.xlsx",
        target_sheet="Ton_NVL",
        target_code_col=1,
        target_code_col_letter="A",
        target_value_col=4,
        target_value_col_letter="D",
        target_start_row=2,
        target_sourcedoc="TARGET-GUID",
        number_convention="strict",
    )


def make_mock_source_bytes(rows: list[tuple[Any, Any]], *, with_header=True) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    start_r = 1
    if with_header:
        ws.cell(1, 2, value="Mã VT")
        ws.cell(1, 13, value="Số lượng tồn")
        start_r = 2

    for idx, (code, qty) in enumerate(rows):
        r = start_r + idx
        ws.cell(r, 2, value=code)
        ws.cell(r, 13, value=qty)

    buf = BytesIO()
    wb.save(buf)
    wb.close()
    return buf.getvalue()


def make_mock_target_bytes(rows: list[tuple[Any, Any]], *, with_header=True, extra_sheets=True) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Ton_NVL"
    start_r = 1
    if with_header:
        ws.cell(1, 1, value="Mã vật tư")
        ws.cell(1, 2, value="Tên vật tư")
        ws.cell(1, 3, value="ĐVT")
        ws.cell(1, 4, value="Tồn kho")
        start_r = 2

    for idx, (code, qty) in enumerate(rows):
        r = start_r + idx
        ws.cell(r, 1, value=code)
        ws.cell(r, 2, value=f"Tên {code}")
        ws.cell(r, 3, value="Cái")
        ws.cell(r, 4, value=qty)

    if extra_sheets:
        ws_extra = wb.create_sheet("DanhMuc")
        ws_extra["A1"] = "Danh mục vật tư"
        ws_extra["B1"] = 12345

    buf = BytesIO()
    wb.save(buf)
    wb.close()
    return buf.getvalue()


def _edit_xml_dimension(workbook_bytes: bytes, sheet_file: str, mode: str) -> bytes:
    """Helper chỉnh sửa XML dimension trong file ZIP để test: missing hoặc understated."""
    in_buf = BytesIO(workbook_bytes)
    out_buf = BytesIO()
    z = zipfile.ZipFile(in_buf, "r")
    dest = zipfile.ZipFile(out_buf, "w")
    try:
        for name in z.namelist():
            data = z.read(name)
            if name == sheet_file:
                root = etree.fromstring(data)
                dims = root.xpath('//*[local-name()="dimension"]')
                if dims:
                    if mode == "missing":
                        dims[0].getparent().remove(dims[0])
                    elif mode == "understated":
                        dims[0].set("ref", "A1:B2")
                data = etree.tostring(root)
            dest.writestr(name, data)
    finally:
        dest.close()
        dest.fp = None
        z.close()
        z.fp = None
    return out_buf.getvalue()


class SyncNVLStockTests(unittest.TestCase):
    def test_normalize_nvl_code(self):
        """Kiểm tra quy tắc chuẩn hóa mã vật tư NVL."""
        # 1. Số nguyên và float nguyên
        self.assertEqual(normalize_nvl_code(12345), "12345")
        self.assertEqual(normalize_nvl_code(12345.0), "12345")
        self.assertEqual(normalize_nvl_code("12345.0"), "12345")
        self.assertEqual(normalize_nvl_code("12345.000"), "12345")

        # 2. Chuỗi có số 0 ở đầu được bảo toàn
        self.assertEqual(normalize_nvl_code("012345"), "012345")
        self.assertEqual(normalize_nvl_code("00987"), "00987")

        # 3. Trim khoảng trắng đầu và cuối, giữ khoảng trắng ở giữa
        self.assertEqual(normalize_nvl_code("  NVL-001  "), "NVL-001")
        self.assertEqual(normalize_nvl_code("  VT 123 ABC  "), "VT 123 ABC")

        # 4. None, boolean, chuỗi rỗng
        self.assertIsNone(normalize_nvl_code(None))
        self.assertIsNone(normalize_nvl_code(""))
        self.assertIsNone(normalize_nvl_code("   "))
        self.assertIsNone(normalize_nvl_code(True))
        self.assertIsNone(normalize_nvl_code(False))

    def test_parse_nvl_quantity(self):
        """Kiểm tra phân tích số lượng tồn kho NVL theo các quy ước và xử lý lỗi."""
        # 1. Số dạng numeric nguyên bản trong Excel
        self.assertEqual(parse_nvl_quantity(0), 0.0)
        self.assertEqual(parse_nvl_quantity(0.0), 0.0)
        self.assertEqual(parse_nvl_quantity(150), 150.0)
        self.assertEqual(parse_nvl_quantity(-25.5), -25.5)
        self.assertAlmostEqual(parse_nvl_quantity(123.456), 123.456)

        # 2. Chuỗi số nguyên và số thập phân rõ ràng
        self.assertEqual(parse_nvl_quantity("0"), 0.0)
        self.assertEqual(parse_nvl_quantity("-50.25"), -50.25)
        self.assertEqual(parse_nvl_quantity("1,5", convention="strict"), 1.5)
        self.assertEqual(parse_nvl_quantity("1.234,56", convention="strict"), 1234.56)
        self.assertEqual(parse_nvl_quantity("1,234.56", convention="strict"), 1234.56)

        # 3. Quy ước Việt Nam ('vi')
        self.assertEqual(parse_nvl_quantity("1,5", convention="vi"), 1.5)
        self.assertEqual(parse_nvl_quantity("1.234,56", convention="vi"), 1234.56)
        self.assertEqual(parse_nvl_quantity("1.234", convention="vi"), 1234.0)
        self.assertEqual(parse_nvl_quantity("1,234", convention="vi"), 1.234)

        # 4. Quy ước Anh ('en')
        self.assertEqual(parse_nvl_quantity("1.5", convention="en"), 1.5)
        self.assertEqual(parse_nvl_quantity("1,234.56", convention="en"), 1234.56)
        self.assertEqual(parse_nvl_quantity("1,234", convention="en"), 1234.0)
        self.assertEqual(parse_nvl_quantity("1.234", convention="en"), 1.234)
        with self.assertRaises(ValueError):
            parse_nvl_quantity("1,5", convention="en")  # Dấu phẩy lẻ không hợp lệ trong EN

        # 5. Chế độ strict: chặn chuỗi mơ hồ khi có đúng 3 số sau dấu phân cách
        with self.assertRaises(ValueError):
            parse_nvl_quantity("1,234", convention="strict")
        with self.assertRaises(ValueError):
            parse_nvl_quantity("1.234", convention="strict")

        # 6. Chặn phân nhóm sai quy cách trong mọi chế độ
        with self.assertRaises(ValueError):
            parse_nvl_quantity("1,2,3")
        with self.assertRaises(ValueError):
            parse_nvl_quantity("1.2.3")
        with self.assertRaises(ValueError):
            parse_nvl_quantity("1..2")
        with self.assertRaises(ValueError):
            parse_nvl_quantity("1,,2")

        # 7. Lỗi chặn: rỗng, boolean, NaN, văn bản không phải số
        with self.assertRaises(ValueError):
            parse_nvl_quantity(None)
        with self.assertRaises(ValueError):
            parse_nvl_quantity("")
        with self.assertRaises(ValueError):
            parse_nvl_quantity("   ")
        with self.assertRaises(ValueError):
            parse_nvl_quantity(True)
        with self.assertRaises(ValueError):
            parse_nvl_quantity(False)
        with self.assertRaises(ValueError):
            parse_nvl_quantity("abc")
        with self.assertRaises(ValueError):
            parse_nvl_quantity(float("nan"))
        with self.assertRaises(ValueError):
            parse_nvl_quantity(float("inf"))

    def test_read_nvl_source_stock_success(self):
        """Đọc nguồn thành công bao gồm tồn 0, âm, lẻ, bỏ qua dòng tổng cộng."""
        cfg = make_mock_config()
        rows = [
            ("VT001", 100),
            ("VT002", 0),
            ("VT003", -15.5),
            ("VT004", 250.75),
            ("Tổng cộng", 335.25),  # Dòng tổng cộng
        ]
        src_bytes = make_mock_source_bytes(rows)
        stock, meta = read_nvl_source_stock(src_bytes, cfg)

        self.assertEqual(len(stock), 4)
        self.assertEqual(stock["VT001"], 100.0)
        self.assertEqual(stock["VT002"], 0.0)
        self.assertEqual(stock["VT003"], -15.5)
        self.assertEqual(stock["VT004"], 250.75)
        self.assertNotIn("Tổng cộng", stock)

    def test_readers_handle_missing_and_understated_dimensions(self):
        """Đọc nguồn và đích đầy đủ ngay cả khi dimension bị thiếu hoặc khai báo hụt (P1)."""
        cfg = make_mock_config()
        src_bytes = make_mock_source_bytes([("VT001", 100), ("VT002", 200)])
        tgt_bytes = make_mock_target_bytes([("VT001", 10), ("VT002", 20)])

        for mode in ["missing", "understated"]:
            # Test nguồn
            src_mod = _edit_xml_dimension(src_bytes, "xl/worksheets/sheet1.xml", mode)
            stock, _ = read_nvl_source_stock(src_mod, cfg)
            self.assertEqual(stock, {"VT001": 100.0, "VT002": 200.0})

            # Test đích
            tgt_mod = _edit_xml_dimension(tgt_bytes, "xl/worksheets/sheet1.xml", mode)
            rec = reconcile_nvl_target(tgt_mod, stock, cfg)
            self.assertEqual(len(rec.changes), 2)
            self.assertEqual(rec.target_codes, {"VT001": 2, "VT002": 3})

    def test_missing_in_source_with_text_and_blank_preserved_safely(self):
        """Bảo toàn ô đích chứa text '-', 'N/A' hoặc blank khi mã thiếu ở nguồn mà không crash float (P2)."""
        cfg = make_mock_config()
        source_stock = {"VT001": 100}
        target_rows = [
            ("VT001", 50),
            ("VT_MISS_1", "-"),
            ("VT_MISS_2", "N/A"),
            ("VT_MISS_3", None),
        ]
        target_bytes = make_mock_target_bytes(target_rows)

        rec = reconcile_nvl_target(target_bytes, source_stock, cfg)
        self.assertEqual(len(rec.changes), 1)
        self.assertEqual(len(rec.missing_in_source), 3)

        patched_bytes = patch_nvl_destination_workbook(target_bytes, rec, cfg)
        verify_res = verify_nvl_patched_workbook(target_bytes, patched_bytes, rec, cfg)
        self.assertTrue(verify_res["ok"])

        # Kiểm tra workbook sau patch
        wb = load_workbook(BytesIO(patched_bytes), data_only=True)
        ws = wb["Ton_NVL"]
        self.assertEqual(ws["D2"].value, 100)  # VT001 cập nhật
        self.assertEqual(ws["D3"].value, "-")  # Giữ nguyên '-'
        self.assertEqual(ws["D4"].value, "N/A")  # Giữ nguyên 'N/A'
        self.assertIsNone(ws["D5"].value)  # Giữ nguyên None
        wb.close()

    def test_unmodified_cells_in_ton_nvl_verified_without_corruption(self):
        """Xác minh kiểm định phát hiện nếu có ô không liên quan trong Ton_NVL bị sửa đổi."""
        cfg = make_mock_config()
        source_stock = {"VT001": 100}
        target_rows = [("VT001", 50), ("VT002", 70)]
        target_bytes = make_mock_target_bytes(target_rows)

        rec = reconcile_nvl_target(target_bytes, source_stock, cfg)
        patched_bytes = patch_nvl_destination_workbook(target_bytes, rec, cfg)

        # Cố tình làm biến đổi cell A2 trong patched XML
        corrupted = BytesIO()
        with zipfile.ZipFile(BytesIO(patched_bytes)) as z, zipfile.ZipFile(corrupted, "w") as dest:
            for item in z.infolist():
                data = z.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(b'VT001', b'VT_HACK')
                dest.writestr(item, data)

        with self.assertRaises(RuntimeError) as ctx:
            verify_nvl_patched_workbook(target_bytes, corrupted.getvalue(), rec, cfg)
        self.assertIn("bị biến đổi cấu trúc XML", str(ctx.exception))

    def test_server_comparison_allows_customxml_and_docprops_change(self):
        """Khi is_server_comparison=True, SharePoint tự cập nhật customXml/* và docProps/* không làm lỗi."""
        cfg = make_mock_config()
        source_stock = {"VT001": 100}
        target_rows = [("VT001", 50)]
        target_bytes = make_mock_target_bytes(target_rows)

        # Thêm customXml/item2.xml và cập nhật docProps/core.xml vào target_bytes ban đầu
        orig_with_custom = BytesIO()
        with zipfile.ZipFile(BytesIO(target_bytes), "r") as z_in, zipfile.ZipFile(orig_with_custom, "w") as z_out:
            for item in z_in.infolist():
                if item.filename != "docProps/core.xml":
                    z_out.writestr(item, z_in.read(item.filename))
            z_out.writestr("customXml/item2.xml", b"<customXml>version_1</customXml>")
            z_out.writestr("docProps/core.xml", b"<core>author_old</core>")
        target_bytes = orig_with_custom.getvalue()

        rec = reconcile_nvl_target(target_bytes, source_stock, cfg)
        patched_bytes = patch_nvl_destination_workbook(target_bytes, rec, cfg)

        # Giả lập SharePoint cập nhật customXml và docProps trên server sau khi upload
        server_sim = BytesIO()
        with zipfile.ZipFile(BytesIO(patched_bytes), "r") as z_in, zipfile.ZipFile(server_sim, "w") as z_out:
            for item in z_in.infolist():
                if item.filename == "customXml/item2.xml":
                    z_out.writestr(item, b"<customXml>version_2_from_sharepoint</customXml>")
                elif item.filename == "docProps/core.xml":
                    z_out.writestr(item, b"<core>author_sharepoint_updated</core>")
                else:
                    z_out.writestr(item, z_in.read(item.filename))
        server_bytes = server_sim.getvalue()

        # Với is_server_comparison=False (mặc định), phải báo lỗi
        with self.assertRaises(RuntimeError) as ctx:
            verify_nvl_patched_workbook(target_bytes, server_bytes, rec, cfg, is_server_comparison=False)
        self.assertTrue(any(k in str(ctx.exception) for k in ("customXml/item2.xml", "docProps/core.xml")))

        # Với is_server_comparison=True, phải bỏ qua metadata SharePoint và kiểm tra hợp lệ
        res = verify_nvl_patched_workbook(target_bytes, server_bytes, rec, cfg, is_server_comparison=True)
        self.assertTrue(res["ok"])

    def test_server_comparison_still_blocks_worksheet_tampering(self):
        """Dù is_server_comparison=True, nếu sheet khác (DanhMuc) bị sửa đổi thì vẫn phải chặn."""
        cfg = make_mock_config()
        source_stock = {"VT001": 100}
        target_rows = [("VT001", 50)]
        target_bytes = make_mock_target_bytes(target_rows)

        rec = reconcile_nvl_target(target_bytes, source_stock, cfg)
        patched_bytes = patch_nvl_destination_workbook(target_bytes, rec, cfg)

        # Cố tình sửa đổi sheet2 (DanhMuc)
        tampered = BytesIO()
        with zipfile.ZipFile(BytesIO(patched_bytes), "r") as z_in, zipfile.ZipFile(tampered, "w") as z_out:
            for item in z_in.infolist():
                data = z_in.read(item.filename)
                if item.filename == "xl/worksheets/sheet2.xml":
                    data = data.replace(b"Danh", b"Hack")
                z_out.writestr(item, data)

        with self.assertRaises(RuntimeError) as ctx:
            verify_nvl_patched_workbook(target_bytes, tampered.getvalue(), rec, cfg, is_server_comparison=True)
        self.assertIn("Phần tử không liên quan 'xl/worksheets/sheet2.xml' trong file ZIP bị thay đổi ngoài ý muốn!", str(ctx.exception))

    def test_target_cell_in_merge_range_blocked(self):
        """Chặn cập nhật khi cột D nằm trong dải ô gộp."""
        cfg = make_mock_config()
        source_stock = {"VT001": 100}

        wb = Workbook()
        ws = wb.active
        ws.title = "Ton_NVL"
        ws["A2"] = "VT001"
        ws["D2"] = 50
        ws.merge_cells("D2:E2")  # Merge cột D và E
        buf = BytesIO()
        wb.save(buf)
        wb.close()

        with self.assertRaises(RuntimeError) as ctx:
            reconcile_nvl_target(buf.getvalue(), source_stock, cfg)
        self.assertIn("gộp", str(ctx.exception).lower())

    def test_target_sheet_protected_blocked(self):
        """Chặn cập nhật khi sheet đích bị bật khóa sheetProtection."""
        cfg = make_mock_config()
        source_stock = {"VT001": 100}
        target_bytes = make_mock_target_bytes([("VT001", 50)])

        # Chèn thẻ sheetProtection vào XML
        prot_bytes = BytesIO()
        with zipfile.ZipFile(BytesIO(target_bytes)) as z, zipfile.ZipFile(prot_bytes, "w") as dest:
            for item in z.infolist():
                data = z.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    root = etree.fromstring(data)
                    ns = etree.QName(root).namespace
                    prot = etree.Element(f"{{{ns}}}sheetProtection", sheet="1")
                    root.append(prot)
                    data = etree.tostring(root)
                dest.writestr(item, data)

        with self.assertRaises(RuntimeError) as ctx:
            reconcile_nvl_target(prot_bytes.getvalue(), source_stock, cfg)
        self.assertIn("sheetprotection", str(ctx.exception).lower())

    def test_read_source_stock_rejects_duplicates(self):
        """Bắt lỗi khi mã vật tư bị lặp trong nguồn."""
        cfg = make_mock_config()
        rows = [
            ("VT001", 100),
            ("VT002", 50),
            ("VT001", 80),  # Lặp mã VT001
        ]
        src_bytes = make_mock_source_bytes(rows)
        with self.assertRaises(RuntimeError) as ctx:
            read_nvl_source_stock(src_bytes, cfg)
        self.assertIn("VT001", str(ctx.exception))
        self.assertIn("lặp", str(ctx.exception).lower())

    def test_read_source_stock_formula_without_cache_rejected(self):
        """Công thức trong cột M không có cached value bị từ chối."""
        cfg = make_mock_config()
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["B2"] = "VT001"
        ws["M2"] = "=SUM(C2:L2)"  # Formula without cached value
        buf = BytesIO()
        wb.save(buf)
        wb.close()

        with self.assertRaises(RuntimeError) as ctx:
            read_nvl_source_stock(buf.getvalue(), cfg)
        self.assertIn("M2", str(ctx.exception))
        self.assertIn("cached", str(ctx.exception).lower())

    def test_reconcile_and_patch_nvl_workbook_full_flow(self):
        """Đối soát và patch hoàn chỉnh: ghép theo mã, thứ tự dòng khác nhau, giữ nguyên mã thiếu."""
        cfg = make_mock_config()

        source_rows = [
            ("VT001", 150),
            ("VT002", 0),
            ("VT003", 75.5),
            ("VT_EXTRA", 999),
        ]
        source_bytes = make_mock_source_bytes(source_rows)
        source_stock, _ = read_nvl_source_stock(source_bytes, cfg)

        target_rows = [
            ("VT003", 50),     # Cần đổi 50 -> 75.5
            ("VT001", 150),    # Không đổi (đã là 150)
            ("VT004", 888),    # Thiếu ở nguồn -> giữ nguyên 888
            ("VT002", 10),     # Cần đổi 10 -> 0
        ]
        target_bytes = make_mock_target_bytes(target_rows)

        rec = reconcile_nvl_target(target_bytes, source_stock, cfg)
        self.assertEqual(rec.status, "completed_with_warnings")
        self.assertEqual(len(rec.changes), 2)  # VT003 và VT002
        self.assertEqual(len(rec.unchanged), 1)  # VT001
        self.assertEqual(len(rec.missing_in_source), 1)  # VT004
        self.assertEqual(rec.source_only, ["VT_EXTRA"])

        # Patch XML
        patched_bytes = patch_nvl_destination_workbook(target_bytes, rec, cfg)
        verify_res = verify_nvl_patched_workbook(target_bytes, patched_bytes, rec, cfg)
        self.assertTrue(verify_res["ok"])

        # Kiểm tra workbook sau patch
        wb = load_workbook(BytesIO(patched_bytes), data_only=True)
        ws = wb["Ton_NVL"]
        self.assertEqual(ws["D2"].value, 75.5)  # VT003 đổi sang 75.5
        self.assertEqual(ws["D3"].value, 150)   # VT001 giữ nguyên 150
        self.assertEqual(ws["D4"].value, 888)   # VT004 thiếu ở nguồn -> giữ nguyên 888
        self.assertEqual(ws["D5"].value, 0)     # VT002 đổi sang 0

        # Kiểm tra sheet phụ DanhMuc hoàn toàn nguyên vẹn
        self.assertIn("DanhMuc", wb.sheetnames)
        self.assertEqual(wb["DanhMuc"]["B1"].value, 12345)
        wb.close()

    def test_reject_formula_in_target_cell(self):
        """Bắt lỗi khi ô đích cột D chứa công thức người dùng."""
        cfg = make_mock_config()
        source_stock = {"VT001": 100}

        wb = Workbook()
        ws = wb.active
        ws.title = "Ton_NVL"
        ws["A2"] = "VT001"
        ws["D2"] = "=VLOOKUP(A2, Kho!A:B, 2, FALSE)"  # Formula in target cell
        buf = BytesIO()
        wb.save(buf)
        wb.close()

        with self.assertRaises(RuntimeError) as ctx:
            reconcile_nvl_target(buf.getvalue(), source_stock, cfg)
        self.assertIn("D2", str(ctx.exception))
        self.assertIn("công thức", str(ctx.exception).lower())

    def test_reject_duplicate_codes_in_target(self):
        """Bắt lỗi khi mã vật tư bị lặp trong đích."""
        cfg = make_mock_config()
        source_stock = {"VT001": 100, "VT002": 200}
        target_rows = [
            ("VT001", 10),
            ("VT002", 20),
            ("VT001", 30),  # Lặp VT001
        ]
        target_bytes = make_mock_target_bytes(target_rows)

        with self.assertRaises(RuntimeError) as ctx:
            reconcile_nvl_target(target_bytes, source_stock, cfg)
        self.assertIn("VT001", str(ctx.exception))
        self.assertIn("lặp", str(ctx.exception).lower())

    def test_no_matching_codes_raises_error(self):
        """Dừng và báo lỗi khi không có mã vật tư nào khớp giữa nguồn và đích."""
        cfg = make_mock_config()
        source_stock = {"VT_SRC_1": 100}
        target_rows = [("VT_TGT_1", 50), ("VT_TGT_2", 60)]
        target_bytes = make_mock_target_bytes(target_rows)

        with self.assertRaises(RuntimeError) as ctx:
            reconcile_nvl_target(target_bytes, source_stock, cfg)
        self.assertIn("không có mã vật tư nào", str(ctx.exception).lower())

    def test_idempotence_second_run_has_zero_changes(self):
        """Chạy lần thứ hai trên file đã patch cho kết quả unchanged, 0 thay đổi."""
        cfg = make_mock_config()
        source_stock = {"VT001": 100.0, "VT002": 200.0}
        target_rows = [("VT001", 50.0), ("VT002", 70.0)]
        target_bytes = make_mock_target_bytes(target_rows)

        rec1 = reconcile_nvl_target(target_bytes, source_stock, cfg)
        self.assertEqual(len(rec1.changes), 2)
        patched_bytes = patch_nvl_destination_workbook(target_bytes, rec1, cfg)

        rec2 = reconcile_nvl_target(patched_bytes, source_stock, cfg)
        self.assertEqual(len(rec2.changes), 0)
        self.assertEqual(len(rec2.unchanged), 2)
        self.assertEqual(rec2.status, "unchanged")

        second_patched_bytes = patch_nvl_destination_workbook(patched_bytes, rec2, cfg)
        self.assertEqual(patched_bytes, second_patched_bytes)

    def test_cli_main_failure_exits_with_code_1_and_writes_error_report(self):
        """CLI main() bắt lỗi, xuất error report và thoát với mã lỗi 1."""
        from sync_nvl_stock import main
        import sys
        from unittest.mock import patch

        with tempfile.TemporaryDirectory() as tmpdir:
            test_args = [
                "sync_nvl_stock.py",
                "--source-file", "nonexistent_source.xlsx",
                "--target-file", "nonexistent_target.xlsx",
                "--out", tmpdir,
            ]
            with patch.object(sys, "argv", test_args):
                with self.assertRaises(SystemExit) as ctx:
                    main()
                self.assertEqual(ctx.exception.code, 1)

            report_file = Path(tmpdir) / "nvl_stock_report.json"
            self.assertTrue(report_file.exists())
            rep = json.loads(report_file.read_text(encoding="utf-8"))
            self.assertEqual(rep["status"], "failed")
            self.assertEqual(rep["phase"], "offline_input")

    def test_cli_main_config_error_writes_error_report_and_exits_1(self):
        """CLI main() gặp lỗi nạp cấu hình vẫn xuất error report và thoát với mã lỗi 1."""
        from sync_nvl_stock import main
        import sys
        from unittest.mock import patch

        with tempfile.TemporaryDirectory() as tmpdir:
            test_args = [
                "sync_nvl_stock.py",
                "--config", "nonexistent_config.json",
                "--out", tmpdir,
            ]
            with patch.object(sys, "argv", test_args):
                with self.assertRaises(SystemExit) as ctx:
                    main()
                self.assertEqual(ctx.exception.code, 1)

            report_file = Path(tmpdir) / "nvl_stock_report.json"
            self.assertTrue(report_file.exists())
            rep = json.loads(report_file.read_text(encoding="utf-8"))
            self.assertEqual(rep["status"], "failed")
            self.assertEqual(rep["phase"], "init_cli")

    def test_preserve_missing_in_source_false_raises_error(self):
        """Khi preserve_missing_in_source=False, gặp mã đích thiếu trong nguồn sẽ dừng và báo lỗi."""
        import io
        import openpyxl
        from sync_nvl_stock import reconcile_nvl_target, NVLConfig
        cfg = NVLConfig(
            source_name="src.xlsm",
            source_path="",
            source_sheet="Sheet1",
            source_code_col=2,
            source_code_col_letter="B",
            source_value_col=13,
            source_value_col_letter="M",
            source_start_row=2,
            source_sourcedoc="",
            target_name="tgt.xlsx",
            target_path="",
            target_sheet="Ton_NVL",
            target_code_col=1,
            target_code_col_letter="A",
            target_value_col=4,
            target_value_col_letter="D",
            target_start_row=2,
            target_sourcedoc="",
            preserve_missing_in_source=False,
        )
        # Target has code 111 and 222
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Ton_NVL"
        ws.cell(1, 1, "Mã NVL")
        ws.cell(1, 4, "Tồn Cuối")
        ws.cell(2, 1, "111")
        ws.cell(2, 4, 10)
        ws.cell(3, 1, "222")
        ws.cell(3, 4, 20)
        buf = io.BytesIO()
        wb.save(buf)
        wb.close()

        # Source only has 111, missing 222
        source_stock = {"111": 15.0}
        with self.assertRaises(RuntimeError) as ctx:
            reconcile_nvl_target(buf.getvalue(), source_stock, cfg)
        self.assertIn("preserve_missing_in_source", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()

