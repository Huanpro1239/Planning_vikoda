"""Điều khiển FIT (SalesUp) bằng Playwright: đăng nhập (CAPTCHA do người nhập), mở
"Báo cáo kế toán > Báo cáo khác > ĐH Bán trong kỳ NPP", đặt bộ lọc và tải file Excel.

Chỉ đọc/xuất báo cáo; không sửa dữ liệu FIT. Phần tử được tìm theo NHÃN tiếng Việt trên
màn hình (không phụ thuộc id HTML), nên giao diện đổi id vẫn chạy; đổi nhãn thì cần sửa.
"""

from __future__ import annotations

import time
from datetime import date
from pathlib import Path
import tempfile
from typing import Callable

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeout, sync_playwright

from .core import MENU_PATH, FitError, ReportSpec, Settings, same_site_url, validate_excel

LOGOUT_TEXT = "Đăng xuất"
BUTTON_TEXT = "Tạo báo cáo"
FORM_TITLE = "Báo cáo đơn hàng bán trong kỳ"

# Chạy trong frame chứa form. Tìm ô theo nhãn hiển thị, đặt giá trị, đánh dấu nút "Tạo báo cáo".
FORM_JS = r"""
(a) => {
  const N = s => (s || '').normalize('NFC').replace(/\s+/g, ' ').trim().toLowerCase();
  const vis = el => !!el && !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
  const own = el => Array.from(el.childNodes).filter(n => n.nodeType === 3).map(n => n.textContent).join(' ');
  const all = () => Array.from(document.querySelectorAll('body *'))
      .filter(el => !['SCRIPT', 'STYLE', 'OPTION', 'SELECT', 'TEXTAREA'].includes(el.tagName));
  const labelOf = t => all().find(el => vis(el) && N(own(el)).replace(/:$/, '') === N(t));
  const after = (lbl, sel) => lbl && Array.from(document.querySelectorAll(sel))
      .find(x => vis(x) && (lbl.compareDocumentPosition(x) & Node.DOCUMENT_POSITION_FOLLOWING));
  const TEXT = 'input:not([type]), input[type="text"], input[type="date"], input[type="search"]';
  const fire = (el, types) => types.forEach(t => el.dispatchEvent(new Event(t, {bubbles: true})));
  const optText = r => {
    if (r.labels && r.labels.length) return N(r.labels[0].textContent);
    let s = '', n = r.nextSibling;
    while (n && !(n.nodeType === 1 && n.tagName === 'INPUT')) { s += n.textContent || ''; if (N(s)) break; n = n.nextSibling; }
    if (N(s)) return N(s);
    const nx = r.parentElement && r.parentElement.nextElementSibling;
    return N(nx ? nx.textContent : '');
  };
  const missing = [], done = {};
  const setDate = (label, value) => {
    const el = after(labelOf(label), TEXT);
    if (!el) { missing.push(label); return; }
    el.focus(); el.value = value; fire(el, ['input', 'change', 'keyup', 'blur']);
    done[label] = el.value;
  };
  const setRadio = (text) => {
    const r = Array.from(document.querySelectorAll('input[type="radio"]')).find(x => vis(x) && optText(x) === N(text));
    if (!r) { missing.push('radio ' + text); return; }
    if (!r.checked) r.click();
    if (!r.checked) { r.checked = true; fire(r, ['change']); }
    done[text] = r.checked;
  };
  const setCheck = (label, want) => {
    const c = after(labelOf(label), 'input[type="checkbox"]');
    if (!c) { missing.push(label); return; }
    if (c.checked !== want) c.click();
    if (c.checked !== want) { c.checked = want; fire(c, ['change']); }
    done[label] = c.checked;
  };
  if (a.inspect) {
    return {labels: ['Từ ngày', 'Đến ngày', 'Mức lấy', 'Không pivot'].map(t => [t, !!labelOf(t)]),
            inputs: Array.from(document.querySelectorAll('input,select,button,textarea')).filter(vis)
              .map(e => [e.tagName, e.type || '', e.id || '', e.name || '', (e.value || '').slice(0, 20)])};
  }
  setDate('Từ ngày', a.tu);
  setDate('Đến ngày', a.den);
  setRadio(a.mucLay);
  setRadio(a.xemThoiGian);
  setCheck('Không pivot', true);
  setCheck('Lấy chi nhánh', false);
  document.querySelectorAll('[data-fit-bot]').forEach(e => e.removeAttribute('data-fit-bot'));
  const want = N(a.button);
  const cands = Array.from(document.querySelectorAll('input[type="button"], input[type="submit"], button, a, span, div, td'))
    .filter(el => vis(el) && (N(el.value) === want || N(own(el)) === want || N(el.textContent) === want));
  // Ưu tiên phần tử có thể bấm, rồi phần tử sâu nhất.
  const rank = el => (['INPUT', 'BUTTON', 'A'].includes(el.tagName) ? 0 : 1);
  cands.sort((x, y) => rank(x) - rank(y) || (x.contains(y) ? 1 : y.contains(x) ? -1 : 0));
  if (cands[0]) cands[0].setAttribute('data-fit-bot', 'btn'); else missing.push(a.button);
  const errBox = after(labelOf('Báo lỗi nhập liệu'), 'textarea, input');
  return {missing, done, error: errBox ? (errBox.value || errBox.textContent || '').trim() : ''};
}
"""

ERROR_JS = r"""
() => {
  const N = s => (s || '').normalize('NFC').replace(/\s+/g, ' ').trim().toLowerCase();
  const lbl = Array.from(document.querySelectorAll('body *')).find(e => N(Array.from(e.childNodes)
      .filter(n => n.nodeType === 3).map(n => n.textContent).join(' ')) === N('Báo lỗi nhập liệu'));
  if (!lbl) return '';
  const box = Array.from(document.querySelectorAll('textarea, input[type="text"]'))
      .find(x => lbl.compareDocumentPosition(x) & Node.DOCUMENT_POSITION_FOLLOWING);
  return box ? (box.value || box.textContent || '').trim() : '';
}
"""

CaptchaProvider = Callable[[bytes, int], "str | None"]


class FitSession:
    def __init__(self, settings: Settings):
        self.s = settings
        self._pw = self._browser = self.context = self.page = None
        self._downloads: list = []
        self._tmp = tempfile.TemporaryDirectory(prefix="fit_")

    # ---- vòng đời ---------------------------------------------------------
    def __enter__(self):
        self._pw = sync_playwright().start()
        try:
            self._browser = self._pw.chromium.launch(headless=self.s.headless)
            self.context = self._browser.new_context(
                accept_downloads=True, ignore_https_errors=self.s.ignore_https_errors,
                locale="vi-VN", timezone_id="Asia/Ho_Chi_Minh", viewport={"width": 1600, "height": 1000},
            )
            self.context.set_default_timeout(self.s.timeout_ms)
            self.context.on("page", self._hook)
            self.page = self.context.new_page()
        except Exception:
            self.close()
            raise
        return self

    def __exit__(self, *exc):
        try:
            self.logout()
        finally:
            self.close()
        return False

    def _hook(self, page):
        page.on("download", lambda d: self._downloads.append(d))

    def close(self):
        for obj in (self.context, self._browser):
            try:
                if obj is not None:
                    obj.close()
            except PlaywrightError:
                pass
        if self._pw is not None:
            self._pw.stop()
        self._pw = self._browser = self.context = self.page = None
        self._tmp.cleanup()

    def screenshot(self) -> bytes | None:
        try:
            return self.page.screenshot(type="png", full_page=False) if self.page else None
        except PlaywrightError:
            return None

    # ---- đăng nhập ----------------------------------------------------------
    def _visible_in_frames(self, text: str) -> bool:
        for frame in self.page.frames:
            try:
                loc = frame.get_by_text(text, exact=True)
                for i in range(min(loc.count(), 5)):
                    if loc.nth(i).is_visible():
                        return True
            except PlaywrightError:
                continue
        return False

    def _login_form_visible(self) -> bool:
        try:
            return self.page.locator("#username").first.is_visible()
        except PlaywrightError:
            return False

    def authenticated(self) -> bool:
        if self.page is None or self.page.url in ("", "about:blank") or self._login_form_visible():
            return False
        return self._visible_in_frames(LOGOUT_TEXT)

    def _page_text(self) -> str:
        try:
            return self.page.inner_text("body", timeout=3000)
        except PlaywrightError:
            return ""

    def login(self, captcha: CaptchaProvider, max_attempts: int = 2):
        self.s.require_fit_credentials()
        try:
            self.page.goto(self.s.base_url, wait_until="domcontentloaded")
        except PlaywrightError as exc:
            if "ERR_CERT" in str(exc) or "SSL" in str(exc):
                raise FitError("Không xác thực được HTTPS của FIT. Đặt biến FIT_IGNORE_HTTPS_ERRORS=true "
                               "nếu máy chủ dùng chứng chỉ tự ký.") from None
            raise FitError(f"Không mở được FIT ({type(exc).__name__}). Máy chạy có thể không truy cập được "
                           f"{self.s.base_url}.") from None
        same_site_url(self.s.base_url, self.page.url)
        if self.authenticated():
            return
        for attempt in range(1, max_attempts + 1):
            self.page.locator("#username").wait_for(state="visible")
            self.page.locator("#username").fill(self.s.username)
            self.page.locator("#password").fill(self.s.password)
            img = self.page.locator('img[src*="CaptchaSvl"]').first
            img.wait_for(state="visible")
            self.page.wait_for_function(
                "() => { const i = document.querySelector('img[src*=\"CaptchaSvl\"]');"
                " return i && i.complete && i.naturalWidth > 0; }")
            code = captcha(img.screenshot(type="png"), attempt)
            if not code:
                raise FitError("Không nhận được mã CAPTCHA (hết giờ hoặc đã hủy). Chưa đăng nhập FIT.")
            self.page.locator("#captchas").fill(code.strip())
            try:
                with self.page.expect_navigation(wait_until="domcontentloaded", timeout=self.s.timeout_ms):
                    self.page.locator('input[type="submit"], button[type="submit"]').first.click()
            except PlaywrightTimeout:
                pass
            for _ in range(20):
                if self.authenticated():
                    print("[FIT] Đăng nhập thành công.")
                    return
                if self._login_form_visible():
                    break
                time.sleep(0.5)
            text = self._page_text().lower()
            if "đã được đăng nhập" in text or "da duoc dang nhap" in text:
                raise FitError("FIT báo tài khoản đang đăng nhập ở nơi khác. Hãy Đăng xuất FIT trên trình duyệt "
                               "của bạn rồi gửi lại lệnh.")
            print(f"[FIT] Đăng nhập chưa thành công (lượt {attempt}/{max_attempts}).")
        raise FitError("FIT từ chối đăng nhập (sai CAPTCHA hoặc sai tài khoản/mật khẩu).")

    def logout(self):
        if self.page is None:
            return
        try:
            if not self.authenticated():
                return
            for frame in self.page.frames:
                loc = frame.get_by_text(LOGOUT_TEXT, exact=True)
                if loc.count() and loc.first.is_visible():
                    loc.first.click()
                    try:
                        self.page.locator("#username").wait_for(state="visible", timeout=15000)
                        print("[FIT] Đã đăng xuất.")
                    except PlaywrightTimeout:
                        print("[FIT] Đã bấm Đăng xuất nhưng chưa thấy màn hình đăng nhập.")
                    return
        except PlaywrightError as exc:
            print(f"[FIT] Không đăng xuất được ({type(exc).__name__}).")

    # ---- mở báo cáo -----------------------------------------------------------
    def _click_text(self, text: str) -> bool:
        for frame in self.page.frames:
            try:
                loc = frame.get_by_text(text, exact=True)
                for i in range(loc.count()):
                    if loc.nth(i).is_visible():
                        loc.nth(i).click()
                        return True
            except PlaywrightError:
                continue
        return False

    def _form_frame(self):
        for frame in self.page.frames:
            try:
                if frame.get_by_text(FORM_TITLE, exact=False).count() and (
                        frame.get_by_text(BUTTON_TEXT, exact=True).count()
                        or frame.locator(f'input[value="{BUTTON_TEXT}"]').count()):
                    return frame
            except PlaywrightError:
                continue
        return None

    def open_report(self):
        if not self.authenticated():
            raise FitError("Phiên FIT chưa đăng nhập hoặc đã hết hạn.")
        if self.s.report_url:
            self.page.goto(same_site_url(self.s.base_url, self.s.report_url), wait_until="domcontentloaded")
        else:
            *parents, leaf = MENU_PATH
            if not self._visible_in_frames(leaf):
                for i, parent in enumerate(parents):
                    # Chỉ bấm mục cha khi mục con chưa hiện (bấm lại sẽ thu gọn menu).
                    child = (list(parents) + [leaf])[i + 1]
                    if not self._visible_in_frames(child):
                        if not self._click_text(parent):
                            raise FitError(f"Không thấy menu “{parent}” trên FIT.")
                        self.page.wait_for_timeout(800)
            if not self._click_text(leaf):
                raise FitError(f"Không thấy menu “{leaf}” trên FIT.")
        deadline = time.monotonic() + self.s.timeout_ms / 1000
        while time.monotonic() < deadline:
            frame = self._form_frame()
            if frame is not None:
                print("[FIT] Đã mở màn hình ĐH Bán trong kỳ NPP.")
                return frame
            self.page.wait_for_timeout(500)
        raise FitError("Đã bấm menu nhưng chưa thấy form “Báo cáo đơn hàng bán trong kỳ”.")

    # ---- xuất ----------------------------------------------------------------
    def export(self, spec: ReportSpec, tu: date, den: date) -> tuple[bytes, str, str]:
        frame = self._form_frame() or self.open_report()
        args = {"tu": tu.strftime(self.s.date_format), "den": den.strftime(self.s.date_format),
                "mucLay": spec.muc_lay, "xemThoiGian": "Ngày hóa đơn", "button": BUTTON_TEXT}
        result = frame.evaluate(FORM_JS, args)
        if result.get("missing"):
            info = frame.evaluate(FORM_JS, {"inspect": True})
            print(f"[FIT] Không thấy trên form: {result['missing']} | nhãn: {info.get('labels')} | "
                  f"ô: {info.get('inputs')[:40]}")
            raise FitError("Không tìm thấy trên form FIT: " + ", ".join(result["missing"]))
        print(f"[FIT] Bộ lọc {spec.key}: {result.get('done')}")
        self._downloads.clear()
        frame.locator('[data-fit-bot="btn"]').first.click()
        deadline = time.monotonic() + self.s.download_timeout_ms / 1000
        while not self._downloads and time.monotonic() < deadline:
            self.page.wait_for_timeout(500)
        if not self._downloads:
            err = ""
            try:
                err = (self._form_frame() or frame).evaluate(ERROR_JS)
            except PlaywrightError:
                pass
            raise FitError(f"{spec.title}: FIT không trả file sau {self.s.download_timeout_ms // 1000}s."
                           + (f" Báo lỗi: {err[:300]}" if err else ""))
        download = self._downloads.pop(0)
        if download.failure():
            raise FitError(f"{spec.title}: tải file thất bại ({download.failure()}).")
        path = Path(self._tmp.name) / f"{spec.key}_{int(time.time())}"
        download.save_as(path)
        content = path.read_bytes()
        path.unlink(missing_ok=True)
        kind = validate_excel(content, spec.title)
        print(f"[FIT] Tải {spec.key}: {download.suggested_filename!r}, {len(content):,} byte ({kind}).")
        return content, download.suggested_filename, kind
