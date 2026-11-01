import io
import threading
import unittest
import zipfile
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlsplit

from openpyxl import Workbook

from fit_ban_hang.core import (REPORTS, FitError, Settings, detect_excel_kind, looks_like_captcha,
                               parse_command, parse_ky, parse_user_ids, target_name)
from fit_ban_hang.telegram import TelegramBot

TODAY = date(2026, 11, 1)


class TestKy(unittest.TestCase):
    def test_mac_dinh_dau_thang_den_hom_nay(self):
        self.assertEqual(parse_ky("", date(2026, 10, 20)), (date(2026, 10, 1), date(2026, 10, 20)))

    def test_thang_truoc(self):
        self.assertEqual(parse_ky("thangtruoc", TODAY), (date(2026, 10, 1), date(2026, 10, 31)))

    def test_ca_thang(self):
        self.assertEqual(parse_ky("9/2026", TODAY), (date(2026, 9, 1), date(2026, 9, 30)))
        self.assertEqual(parse_ky("2026-02", TODAY), (date(2026, 2, 1), date(2026, 2, 28)))
        self.assertEqual(parse_ky("11/2026", TODAY), (date(2026, 11, 1), date(2026, 11, 1)))

    def test_khoang_ngay(self):
        self.assertEqual(parse_ky("01/10/2026 15/10/2026", TODAY), (date(2026, 10, 1), date(2026, 10, 15)))
        self.assertEqual(parse_ky("2026-10-01 -> 2026-10-15", TODAY), (date(2026, 10, 1), date(2026, 10, 15)))
        self.assertEqual(parse_ky("5/10/2026", TODAY), (date(2026, 10, 5), date(2026, 10, 5)))

    def test_loi(self):
        for bad in ("15/10/2026 01/10/2026", "01/12/2026", "12/2026", "abc", "1/1/2026 2/2/2026 3/3/2026"):
            with self.assertRaises(FitError, msg=bad):
                parse_ky(bad, TODAY)


class TestLenh(unittest.TestCase):
    def test_parse_command(self):
        self.assertEqual(parse_command("/banhang 10/2026"), ("banhang", "10/2026"))
        self.assertEqual(parse_command("/BanHang@VikodaBot"), ("banhang", ""))
        self.assertIsNone(parse_command("AB12"))

    def test_captcha_va_id(self):
        self.assertTrue(looks_like_captcha(" aB3x9 "))
        self.assertFalse(looks_like_captcha("ab"))
        self.assertFalse(looks_like_captcha("ma la ab12"))
        self.assertEqual(parse_user_ids("123456, 7891011"), frozenset({123456, 7891011}))
        with self.assertRaises(FitError):
            parse_user_ids("abc")


def _xlsx_bytes():
    wb = Workbook()
    wb.active["A1"] = "x"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestExcel(unittest.TestCase):
    def test_nhan_dien(self):
        self.assertEqual(detect_excel_kind(_xlsx_bytes()), "ooxml")
        self.assertEqual(detect_excel_kind(bytes.fromhex("D0CF11E0A1B11AE1") + b"\0" * 100), "xls")
        self.assertIsNone(detect_excel_kind(b"<html><input id=password></html>"))
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("a.txt", "x")
        self.assertIsNone(detect_excel_kind(buf.getvalue()))

    def test_ten_file_theo_anh_3(self):
        self.assertEqual(target_name(REPORTS[0], "BCDonHangBanTrongKyNPP.xlsm", "ooxml"), "BCDonHangBanTrongKyNPP.xlsm")
        self.assertEqual(target_name(REPORTS[1], "BCDonHangBanTrongKyNPP.xlsm", "ooxml"), "Lay So Hoa don.xlsm")
        self.assertEqual(target_name(REPORTS[1], "x.xlsx", "ooxml"), "Lay So Hoa don.xlsx")
        self.assertEqual(target_name(REPORTS[0], "x.xls", "xls"), "BCDonHangBanTrongKyNPP.xls")


class FakeSession:
    def __init__(self, batches):
        self.batches = list(batches)
        self.sent = []

    def post(self, url, data=None, files=None, timeout=None):
        method = url.rsplit("/", 1)[-1]
        self.sent.append((method, data))
        result = self.batches.pop(0) if method == "getUpdates" and self.batches else ([] if method == "getUpdates" else {"message_id": 99})

        class R:
            status_code = 200

            def json(self_inner):
                return {"ok": True, "result": result}
        return R()


def _upd(uid, text, user=111, reply_to=None, chat_type="private"):
    msg = {"message_id": uid, "chat": {"id": user, "type": chat_type}, "from": {"id": user}, "text": text}
    if reply_to:
        msg["reply_to_message"] = {"message_id": reply_to}
    return {"update_id": uid, "message": msg}


class TestTelegram(unittest.TestCase):
    def test_chi_nhan_ma_tu_nguoi_duoc_phep(self):
        fake = FakeSession([[_upd(1, "AB12", user=222), _upd(2, "AB12", chat_type="group"),
                             _upd(3, "xin chao ban"), _upd(4, "XY99", reply_to=50)]])
        bot = TelegramBot("t", {111}, session=fake)
        self.assertEqual(bot.wait_captcha(111, 50, 30), "XY99")
        self.assertEqual(bot.offset, 5)
        # tin "xin chao ban" được nhắc chờ mã
        self.assertTrue(any(m == "sendMessage" for m, _ in fake.sent))

    def test_huy(self):
        bot = TelegramBot("t", {111}, session=FakeSession([[_upd(1, "/huy")]]))
        self.assertIsNone(bot.wait_captcha(111, 50, 30))


# ----------------------------------------------------------------------------------
# Kiểm thử đầu-cuối với FIT mô phỏng (bố cục giống ảnh: menu trái, form trong iframe).
LOGIN = """<html><body><form method=post action="/FIT/login">
<input id=username name=username><input id=password name=password type=password>
<img src="/FIT/CaptchaSvl"><input id=captchas name=captchas><input type=submit value="Đăng nhập">
</form>%s</body></html>"""
MAIN = """<html><body>
<div style="float:right"><a href="/FIT/logout">Đăng xuất</a> <a href="#">Đổi mật khẩu</a></div>
<div id=menu style="float:left;width:250px">
 <div onclick="t('m1')">Báo cáo kế toán</div>
 <div id=m1 style="display:none">
   <div onclick="t('m2')">Báo cáo khác</div>
   <div id=m2 style="display:none">
     <div onclick="document.getElementById('f').src='/FIT/Center/report.jsp'">ĐH Bán trong kỳ NPP</div>
     <div>Công nợ vỏ két</div>
   </div>
 </div>
 <div>Báo cáo quản trị</div>
</div>
<iframe id=f style="width:1200px;height:700px" src="about:blank"></iframe>
<script>function t(id){var e=document.getElementById(id);e.style.display=e.style.display=='none'?'block':'none';}</script>
</body></html>"""
REPORT = """<html><body>
<fieldset><legend>Báo lỗi nhập liệu</legend><textarea id=err></textarea></fieldset>
<fieldset><legend>Báo cáo đơn hàng bán trong kỳ</legend><table>
<tr><td><b>Từ ngày</b></td><td><input type=text id=tungay></td><td>Đến ngày</td><td><input type=text id=denngay></td></tr>
<tr><td>Miền</td><td><input type=text id=mien></td><td>Khu Vực</td><td><input type=text id=kv></td></tr>
<tr><td>Kho</td><td><select id=kho><option>All</option></select></td></tr>
<tr><td>Mức lấy</td><td><input type=radio name=muclay value=0> Hóa Đơn <input type=radio name=muclay value=1 checked> Đơn hàng</td>
    <td>Xem thời gian</td><td><input type=radio name=tg value=0 checked>Ngày hóa đơn <input type=radio name=tg value=1>Ngày ghi sổ</td></tr>
<tr><td>Lấy chi nhánh</td><td><input type=checkbox id=cn checked></td></tr>
<tr><td>Không pivot</td><td><input type=checkbox id=pv></td></tr>
</table>
<div class=btn onclick="go()"><span>Tạo báo cáo</span></div></fieldset>
<script>
function go(){
  var tu=document.getElementById('tungay').value, den=document.getElementById('denngay').value;
  if(!tu||!den){document.getElementById('err').value='Chưa nhập ngày';return;}
  var ml=document.querySelector('input[name=muclay]:checked').value;
  var q='tu='+encodeURIComponent(tu)+'&den='+encodeURIComponent(den)+'&ml='+ml
    +'&pv='+document.getElementById('pv').checked+'&cn='+document.getElementById('cn').checked
    +'&tg='+document.querySelector('input[name=tg]:checked').value;
  setTimeout(function(){window.location='/FIT/export?'+q;},300);
}
</script></body></html>"""


class FakeFit(BaseHTTPRequestHandler):
    state = {"logged": False, "exports": [], "logins": 0}

    def log_message(self, *a):
        pass

    def _send(self, body, ctype="text/html; charset=utf-8", extra=None):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        st = self.state
        path, q = urlsplit(self.path).path, parse_qs(urlsplit(self.path).query)
        if path == "/FIT/CaptchaSvl":
            png = bytes.fromhex("89504E470D0A1A0A0000000D49484452000000010000000108060000001F15C4890000000D49444154789C6360000002000154A24F5D0000000049454E44AE426082")
            return self._send(png, "image/png")
        if path == "/FIT/logout":
            st["logged"] = False
            return self._send(LOGIN % "")
        if path == "/FIT/Center/report.jsp" and st["logged"]:
            return self._send(REPORT)
        if path == "/FIT/export" and st["logged"]:
            st["exports"].append({k: v[0] for k, v in q.items()})
            wb = Workbook()
            wb.active["A1"] = "ml=" + q["ml"][0]
            buf = io.BytesIO()
            wb.save(buf)
            return self._send(buf.getvalue(), "application/vnd.ms-excel.sheet.macroEnabled.12",
                              {"Content-Disposition": 'attachment; filename="BCDonHangBanTrongKyNPP.xlsm"'})
        return self._send(MAIN if st["logged"] else LOGIN % "")

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        form = parse_qs(self.rfile.read(length).decode())
        self.state["logins"] += 1
        ok = form.get("username") == ["u"] and form.get("password") == ["p"] and form.get("captchas") == ["DEMO"]
        self.state["logged"] = ok
        self.send_response(302)
        self.send_header("Location", "/FIT/Center/mainpage.jsp" if ok else "/FIT/")
        self.end_headers()


def _chromium_available():
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            pw.chromium.launch().close()
        return True
    except Exception:
        return False


@unittest.skipUnless(_chromium_available(), "Chưa cài Chromium cho Playwright")
class TestFitDauCuoi(unittest.TestCase):
    def setUp(self):
        FakeFit.state = {"logged": False, "exports": [], "logins": 0}
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), FakeFit)
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        self.settings = Settings(base_url=f"http://127.0.0.1:{self.server.server_port}/FIT/",
                                 username="u", password="p", timeout_ms=10000, download_timeout_ms=10000)

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()

    def test_dang_nhap_loc_va_tai_2_file(self):
        from fit_ban_hang.fit import FitSession

        codes = iter(["SAI1", "DEMO"])
        with FitSession(self.settings) as fit:
            fit.login(lambda png, attempt: next(codes))
            fit.open_report()
            out = [fit.export(spec, date(2026, 10, 1), date(2026, 10, 31)) for spec in REPORTS]
        self.assertEqual(FakeFit.state["logins"], 2)
        self.assertFalse(FakeFit.state["logged"], "phải đăng xuất sau khi xong")
        ex = FakeFit.state["exports"]
        self.assertEqual([e["ml"] for e in ex], ["1", "0"])  # Đơn hàng rồi Hóa Đơn
        for e in ex:
            self.assertEqual((e["tu"], e["den"], e["pv"], e["cn"], e["tg"]), ("01/10/2026", "31/10/2026", "true", "false", "0"))
        names = [target_name(spec, sug, kind) for spec, (c, sug, kind) in zip(REPORTS, out)]
        self.assertEqual(names, ["BCDonHangBanTrongKyNPP.xlsm", "Lay So Hoa don.xlsm"])

    def test_het_ma_captcha(self):
        from fit_ban_hang.fit import FitSession

        with FitSession(self.settings) as fit, self.assertRaises(FitError):
            fit.login(lambda png, attempt: None)
        self.assertEqual(FakeFit.state["logins"], 0)

    def test_run_job_qua_telegram_dry_run(self):
        import os
        import tempfile
        from fit_ban_hang.__main__ import run_job

        class Bot:
            def __init__(self):
                self.msgs, self.docs = [], []

            def send_message(self, chat, text):
                self.msgs.append(text)

            def send_photo(self, chat, png, caption, force_reply=False):
                return 7

            def wait_captcha(self, chat, msg_id, wait):
                return "DEMO"

            def send_document(self, chat, name, content, caption=""):
                self.docs.append(name)

        bot, cwd = Bot(), os.getcwd()
        with tempfile.TemporaryDirectory() as tmp:
            os.chdir(tmp)
            try:
                rc = run_job(self.settings, date(2026, 10, 1), date(2026, 10, 31), bot=bot, chat_id=1, dry_run=True)
                self.assertTrue(os.path.exists("Lay So Hoa don.xlsm"))
            finally:
                os.chdir(cwd)
        self.assertEqual(rc, 0)
        self.assertEqual(bot.docs, ["BCDonHangBanTrongKyNPP.xlsm", "Lay So Hoa don.xlsm"])
        self.assertTrue(bot.msgs[-1].startswith("✅"))


if __name__ == "__main__":
    unittest.main()
