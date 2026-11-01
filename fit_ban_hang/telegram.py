"""Client Telegram Bot API tối giản (requests, long polling). Không log token/nội dung tin nhắn."""

from __future__ import annotations

import time

import requests

from .core import FitError, looks_like_captcha, parse_command


class TelegramBot:
    def __init__(self, token: str, allowed_user_ids, *, session=None):
        self._base = f"https://api.telegram.org/bot{token}/"
        self.allowed = frozenset(allowed_user_ids)
        self.http = session or requests.Session()
        self.offset: int | None = None

    # ---- Bot API -------------------------------------------------------
    def _call(self, method, *, data=None, files=None, timeout=60):
        try:
            response = self.http.post(self._base + method, data=data, files=files, timeout=timeout)
        except requests.RequestException as exc:
            # Không in exc: URL có chứa token.
            raise FitError(f"Không kết nối được Telegram ({type(exc).__name__}).") from None
        try:
            payload = response.json()
        except ValueError:
            payload = {}
        if response.status_code == 409:
            return None  # một tiến trình khác đang getUpdates cùng token
        if not payload.get("ok"):
            raise FitError(f"Telegram {method} lỗi {response.status_code}: {payload.get('description', '')}")
        return payload.get("result")

    def get_updates(self, timeout: int = 25) -> list[dict]:
        data = {"timeout": max(0, int(timeout)), "allowed_updates": '["message"]'}
        if self.offset is not None:
            data["offset"] = self.offset
        result = self._call("getUpdates", data=data, timeout=timeout + 15)
        if result is None:
            print("[Telegram] Token đang được tiến trình khác dùng (409) – chờ rồi thử lại.")
            time.sleep(min(10, max(1, timeout)))
            return []
        for upd in result:
            self.offset = max(self.offset or 0, int(upd["update_id"]) + 1)
        return result

    def confirm(self):
        """Báo Telegram đã xử lý xong các update đã đọc (tiến trình sau không nhận lại)."""
        if self.offset is not None:
            try:
                self._call("getUpdates", data={"offset": self.offset, "timeout": 0, "limit": 1}, timeout=20)
            except FitError:
                pass

    def send_message(self, chat_id: int, text: str) -> int | None:
        res = self._call("sendMessage", data={"chat_id": chat_id, "text": text[:4000],
                                              "disable_web_page_preview": "true"})
        return (res or {}).get("message_id")

    def send_photo(self, chat_id: int, png: bytes, caption: str, *, force_reply: bool = False) -> int | None:
        data = {"chat_id": chat_id, "caption": caption[:1000]}
        if force_reply:
            data["reply_markup"] = '{"force_reply":true,"input_field_placeholder":"Mã CAPTCHA"}'
        res = self._call("sendPhoto", data=data, files={"photo": ("fit.png", png, "image/png")}, timeout=90)
        return (res or {}).get("message_id")

    def send_document(self, chat_id: int, filename: str, content: bytes, caption: str = "") -> int | None:
        res = self._call("sendDocument", data={"chat_id": chat_id, "caption": caption[:1000]},
                         files={"document": (filename, content, "application/octet-stream")}, timeout=180)
        return (res or {}).get("message_id")

    # ---- Quyền & tin nhắn ------------------------------------------------
    def authorized_message(self, update: dict) -> dict | None:
        msg = update.get("message") or {}
        chat, user = msg.get("chat") or {}, msg.get("from") or {}
        if chat.get("type") == "private" and user.get("id") in self.allowed and isinstance(msg.get("text"), str):
            return msg
        return None

    def wait_captcha(self, chat_id: int, photo_message_id: int | None, wait_s: int) -> str | None:
        """Chờ người dùng được phép trả lời mã CAPTCHA. None nếu hết giờ hoặc /huy."""
        deadline = time.monotonic() + wait_s
        while True:
            remaining = int(deadline - time.monotonic())
            if remaining <= 0:
                return None
            for upd in self.get_updates(timeout=min(25, remaining)):
                msg = self.authorized_message(upd)
                if msg is None:
                    continue
                text = msg["text"].strip()
                cmd = parse_command(text)
                if cmd and cmd[0] in ("huy", "cancel"):
                    self.send_message(msg["chat"]["id"], "Đã hủy lượt lấy báo cáo.")
                    return None
                if cmd and cmd[0] == "captcha" and looks_like_captcha(cmd[1]):
                    return cmd[1].strip()
                reply_id = (msg.get("reply_to_message") or {}).get("message_id")
                if not cmd and looks_like_captcha(text) and (reply_id in (None, photo_message_id)):
                    return text
                self.send_message(msg["chat"]["id"], "Đang chờ mã CAPTCHA của ảnh vừa gửi (hoặc /huy để dừng).")
