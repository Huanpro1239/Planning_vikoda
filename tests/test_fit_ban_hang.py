import io
import os
import tempfile
import unittest
import zipfile
from datetime import date

from openpyxl import Workbook

from fit_ban_hang.__main__ import run
from fit_ban_hang.core import (REPORTS, FitError, Settings, detect_excel_kind, ooxml_extension, parse_ky,
                               target_name)
from fit_ban_hang.worker import WorkerClient

TODAY = date(2026, 11, 1)


def _xlsx(macro=False):
    wb = Workbook()
    wb.active["A1"] = "x"
    buf = io.BytesIO()
    wb.save(buf)
    data = buf.getvalue()
    if not macro:
        return data
    src, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w") as dst:
        for item in src.infolist():
            body = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                body = body.replace(b"spreadsheetml.sheet.main+xml", b"ms-excel.sheet.macroEnabled.main+xml")
            dst.writestr(item, body)
    return out.getvalue()


class TestKy(unittest.TestCase):
    def test_cac_dang_ky(self):
        self.assertEqual(parse_ky("", date(2026, 10, 20)), (date(2026, 10, 1), date(2026, 10, 20)))
        self.assertEqual(parse_ky("thangtruoc", TODAY), (date(2026, 10, 1), date(2026, 10, 31)))
        self.assertEqual(parse_ky("9/2026", TODAY), (date(2026, 9, 1), date(2026, 9, 30)))
        self.assertEqual(parse_ky("2026-10-01 2026-10-15", TODAY), (date(2026, 10, 1), date(2026, 10, 15)))
        self.assertEqual(parse_ky("01/10/2026 15/10/2026", TODAY), (date(2026, 10, 1), date(2026, 10, 15)))

    def test_loi(self):
        for bad in ("15/10/2026 01/10/2026", "01/12/2026", "12/2026", "abc", "1/1/2026 2/2/2026 3/3/2026"):
            with self.assertRaises(FitError, msg=bad):
                parse_ky(bad, TODAY)


class TestExcel(unittest.TestCase):
    def test_nhan_dien_va_ten_theo_anh_3(self):
        self.assertEqual(detect_excel_kind(_xlsx()), "ooxml")
        self.assertIsNone(detect_excel_kind(b"<html><form name=loginForm></html>"))
        self.assertEqual(ooxml_extension(_xlsx(macro=True)), ".xlsm")
        self.assertEqual(ooxml_extension(_xlsx()), ".xlsx")
        self.assertEqual(target_name(REPORTS[0], _xlsx(True), "ooxml"), "BCDonHangBanTrongKyNPP.xlsm")
        self.assertEqual(target_name(REPORTS[1], _xlsx(True), "ooxml"), "Lay So Hoa don.xlsm")
        self.assertEqual(target_name(REPORTS[1], b"", "xls"), "Lay So Hoa don.xls")


class Resp:
    def __init__(self, status, content=b"", js=None):
        self.status_code, self.content, self._js = status, content, js

    def json(self):
        if self._js is None:
            raise ValueError
        return self._js


class FakeHttp:
    def __init__(self, responses):
        self.responses, self.calls = list(responses), []

    def get(self, url, params=None, headers=None, timeout=None, allow_redirects=True):
        self.calls.append((url, dict(params), dict(headers)))
        return self.responses.pop(0)


SETTINGS = Settings(cf_client_id="id.access", cf_client_secret="sec")


class TestWorkerClient(unittest.TestCase):
    def test_goi_dung_route_va_header_access(self):
        http = FakeHttp([Resp(409, js={"code": "ERP_BUSY"}), Resp(200, _xlsx(True))])
        content, kind = WorkerClient(SETTINGS, session=http, sleep=lambda s: None).fetch(
            REPORTS[1], date(2026, 10, 1), date(2026, 10, 31))
        self.assertEqual(kind, "ooxml")
        url, params, headers = http.calls[-1]
        self.assertEqual(url, "https://erp-nhan-hang.nguyensihuan711994.workers.dev/api/integrations/ban-hang-npp/file")
        self.assertEqual(params, {"start": "2026-10-01", "end": "2026-10-31", "muc": "hoa_don"})
        self.assertEqual((headers["CF-Access-Client-Id"], headers["CF-Access-Client-Secret"]), ("id.access", "sec"))
        self.assertEqual(len(http.calls), 2)

    def test_loi_ro_rang(self):
        cases = [
            (Resp(401, js={"code": "ERP_LOGIN_REQUIRED", "error": "x"}), "/dangnhap"),
            (Resp(302), "Cloudflare Access"),
            (Resp(200, b"<html>login</html>"), "không phải Excel"),
            (Resp(422, js={"code": "REPORT_NOT_EXPORTED", "error": "Chưa chọn ngày"}), "Chưa chọn ngày"),
        ]
        for resp, text in cases:
            with self.assertRaises(FitError) as ctx:
                WorkerClient(SETTINGS, session=FakeHttp([resp]), sleep=lambda s: None).fetch(
                    REPORTS[0], date(2026, 10, 1), date(2026, 10, 2))
            self.assertIn(text, str(ctx.exception))
        with self.assertRaises(FitError):
            WorkerClient(Settings(), session=FakeHttp([]))


class FakeClient:
    def __init__(self, fail=None):
        self.fail, self.seen = fail, []

    def fetch(self, spec, tu, den):
        self.seen.append((spec.key, tu, den))
        if self.fail:
            raise FitError(self.fail)
        return _xlsx(True), "ooxml"


class FakeNotifier:
    def __init__(self):
        self.msgs, self.docs = [], []

    def send_message(self, text):
        self.msgs.append(text)

    def send_document(self, name, content, caption=""):
        self.docs.append(name)


class TestRun(unittest.TestCase):
    def test_dry_run_hai_file_va_bao_telegram(self):
        client, bot, cwd = FakeClient(), FakeNotifier(), os.getcwd()
        with tempfile.TemporaryDirectory() as tmp:
            os.chdir(tmp)
            try:
                rc = run(SETTINGS, date(2026, 10, 1), date(2026, 10, 31), dry_run=True, client=client, notifier=bot)
                self.assertTrue(os.path.exists("BCDonHangBanTrongKyNPP.xlsm"))
                self.assertTrue(os.path.exists("Lay So Hoa don.xlsm"))
            finally:
                os.chdir(cwd)
        self.assertEqual(rc, 0)
        self.assertEqual([s[0] for s in client.seen], ["don_hang", "hoa_don"])
        self.assertEqual(bot.docs, ["BCDonHangBanTrongKyNPP.xlsm", "Lay So Hoa don.xlsm"])
        self.assertTrue(bot.msgs[-1].startswith("✅"))

    def test_loi_khong_ghi_sharepoint(self):
        bot = FakeNotifier()

        def no_graph():
            raise AssertionError("không được ghi SharePoint khi chưa lấy đủ file")

        rc = run(SETTINGS, date(2026, 10, 1), date(2026, 10, 31), client=FakeClient(fail="Phiên FIT đã hết hạn"),
                 notifier=bot, graph_factory=no_graph)
        self.assertEqual(rc, 1)
        self.assertIn("Phiên FIT đã hết hạn", bot.msgs[-1])


if __name__ == "__main__":
    unittest.main()
