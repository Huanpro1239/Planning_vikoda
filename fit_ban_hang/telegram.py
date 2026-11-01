"""Gửi thông báo/file vào Telegram (chỉ gửi, KHÔNG getUpdates).

Bot do Worker erp-nhan-hang đọc tin nhắn bằng getUpdates; GitHub chỉ được gửi, nếu cũng
đọc tin nhắn sẽ tranh với Worker (lỗi 409). Không log token/nội dung.
"""

from __future__ import annotations

import requests

from .core import FitError


class TelegramNotifier:
    def __init__(self, token: str, chat_id: str, *, session=None):
        self._base = f"https://api.telegram.org/bot{token}/"
        self.chat_id = chat_id
        self.http = session or requests.Session()

    def _call(self, method, *, data=None, files=None, timeout=60):
        try:
            response = self.http.post(self._base + method, data=data, files=files, timeout=timeout)
        except requests.RequestException as exc:
            raise FitError(f"Không kết nối được Telegram ({type(exc).__name__}).") from None
        try:
            payload = response.json()
        except ValueError:
            payload = {}
        if not payload.get("ok"):
            raise FitError(f"Telegram {method} lỗi {response.status_code}: {payload.get('description', '')}")
        return payload.get("result")

    def send_message(self, text: str):
        return self._call("sendMessage", data={"chat_id": self.chat_id, "text": text[:4000],
                                               "disable_web_page_preview": "true"})

    def send_document(self, filename: str, content: bytes, caption: str = ""):
        return self._call("sendDocument", data={"chat_id": self.chat_id, "caption": caption[:1000]},
                          files={"document": (filename, content, "application/octet-stream")}, timeout=180)
