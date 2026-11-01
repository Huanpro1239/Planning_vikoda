"""Gọi Worker erp-nhan-hang lấy file Excel gốc của FIT "ĐH Bán trong kỳ NPP".

Worker giữ phiên FIT (tự gia hạn 15 phút/lần, CAPTCHA qua Telegram khi hết hạn), nên GitHub
không đăng nhập FIT và không cần mật khẩu FIT. Xác thực bằng service token Cloudflare Access
(cùng token máy đang dùng cho Nhận hàng); Worker chỉ cho token này đọc đúng route này.
"""

from __future__ import annotations

import time
from datetime import date

import requests

from .core import WORKER_PATH, FitError, ReportSpec, Settings, validate_excel


class WorkerClient:
    def __init__(self, settings: Settings, *, session=None, sleep=time.sleep):
        settings.require_worker_credentials()
        self.s = settings
        self.http = session or requests.Session()
        self.sleep = sleep

    def fetch(self, spec: ReportSpec, tu: date, den: date, *, attempts: int = 12) -> tuple[bytes, str]:
        url = self.s.worker_origin + WORKER_PATH
        params = {"start": tu.isoformat(), "end": den.isoformat(), "muc": spec.key}
        headers = {"CF-Access-Client-Id": self.s.cf_client_id,
                   "CF-Access-Client-Secret": self.s.cf_client_secret,
                   "User-Agent": "Planning_vikoda/fit_ban_hang"}
        for i in range(1, attempts + 1):
            try:
                resp = self.http.get(url, params=params, headers=headers, timeout=150, allow_redirects=False)
            except requests.RequestException as exc:
                if i == attempts:
                    raise FitError(f"Không kết nối được Worker ({type(exc).__name__}).") from None
                self.sleep(20)
                continue
            if resp.status_code == 200:
                content = resp.content
                kind = validate_excel(content, spec.title)
                print(f"[Worker] {spec.key}: {len(content):,} byte ({kind}).")
                return content, kind
            info = _json(resp)
            code = info.get("code", "")
            if resp.status_code == 409 or code == "ERP_BUSY":
                print(f"[Worker] FIT đang bận (lượt {i}/{attempts}), chờ 30s.")
                if i < attempts:
                    self.sleep(30)
                    continue
                raise FitError("FIT đang bận xuất dữ liệu khác quá lâu. Thử lại sau.")
            if resp.status_code in (301, 302, 303, 307, 308, 403) and not info:
                raise FitError("Cloudflare Access từ chối GitHub: kiểm tra secret FIT_CF_ACCESS_CLIENT_ID/SECRET "
                               "và đã deploy Worker bản có /banhang.")
            if resp.status_code == 401 or code == "ERP_LOGIN_REQUIRED":
                raise FitError("Phiên FIT đã hết hạn. Gửi /dangnhap cho bot, trả lời mã bảo vệ rồi gửi lại /banhang.")
            if resp.status_code in (502, 503, 504) and i < attempts:
                print(f"[Worker] HTTP {resp.status_code} (lượt {i}/{attempts}), thử lại sau 20s.")
                self.sleep(20)
                continue
            raise FitError(f"{spec.title}: Worker trả HTTP {resp.status_code}"
                           + (f" – {info.get('error')}" if info.get("error") else "."))
        raise FitError("Không lấy được báo cáo từ Worker.")


def _json(resp) -> dict:
    try:
        data = resp.json()
        return data if isinstance(data, dict) else {}
    except ValueError:
        return {}
