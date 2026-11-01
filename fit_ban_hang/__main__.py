"""FIT "ĐH Bán trong kỳ NPP" -> SharePoint, điều khiển qua Telegram.

    python -X utf8 -m fit_ban_hang run [--ky "10/2026"] [--dry-run]   # chạy một lượt
    python -X utf8 -m fit_ban_hang listen --seconds 900                 # nghe lệnh Telegram (dùng trong auto-sync)

Lệnh Telegram (chat riêng, chỉ ID trong TELEGRAM_ALLOWED_USER_IDS):
    /banhang                 đầu tháng -> hôm nay
    /banhang thangtruoc      cả tháng trước
    /banhang 10/2026         cả tháng 10/2026
    /banhang 01/10/2026 15/10/2026
    /id                      xem ID Telegram của bạn
Bot gửi ảnh CAPTCHA -> trả lời ảnh bằng mã -> bot tải 2 file, ghi SharePoint và gửi file vào chat.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from .core import REPORTS, VN_TZ, FitError, Settings, fmt_ky, now_vn, parse_command, parse_ky, target_name

HELP = (
    "Bot lấy báo cáo FIT → SharePoint\n"
    "/banhang — ĐH bán trong kỳ NPP từ đầu tháng đến hôm nay\n"
    "/banhang thangtruoc — cả tháng trước\n"
    "/banhang 10/2026 — cả tháng\n"
    "/banhang 01/10/2026 15/10/2026 — khoảng ngày\n"
    "/huy — dừng khi đang chờ CAPTCHA\n"
    "/id — xem ID Telegram\n\n"
    "Mỗi lượt lấy 2 file: BCDonHangBanTrongKyNPP (Mức lấy Đơn hàng) và Lay So Hoa don (Mức lấy Hóa đơn), "
    "ghi đè trên SharePoint Planning."
)
XLSM_CT = "application/vnd.ms-excel.sheet.macroEnabled.12"
STATE_PATH = Path(os.environ.get("RUNNER_TEMP") or ".") / ".fit_ban_hang_state.json"


# ---- SharePoint ----------------------------------------------------------------
def _graph():
    from sharepoint.client import GraphClient, get_access_token, retry_graph_read

    graph = GraphClient(get_access_token())
    site_id = retry_graph_read("site", graph.get_site_id, attempts=6, max_delay_seconds=30)
    drive_id = retry_graph_read("drive", lambda: graph.get_default_drive_id(site_id), attempts=6,
                                max_delay_seconds=30)
    return graph, drive_id


def upload(graph, drive_id: str, path: str, content: bytes, attempts: int = 5) -> dict:
    from sharepoint.client import GRAPH, GraphRequestError

    url = f"{GRAPH}/drives/{drive_id}/root:/{quote(path.strip('/'), safe='/')}:/content"
    ctype = XLSM_CT if path.lower().endswith(".xlsm") else "application/octet-stream"
    for i in range(1, attempts + 1):
        try:
            resp = graph.session.put(url, params={"@microsoft.graph.conflictBehavior": "replace"},
                                     headers={"Content-Type": ctype}, data=content, timeout=180)
            graph._raise(resp)
            return resp.json()
        except GraphRequestError as exc:
            if i == attempts or exc.status_code not in (409, 412, 423, 429, 500, 502, 503, 504):
                if exc.status_code == 423:
                    raise FitError(f"File {path.rsplit('/', 1)[-1]} đang bị khóa (có người mở trên Excel desktop).") from None
                raise
            print(f"[SharePoint] Ghi lỗi {exc.status_code}, thử lại sau {15 * i}s")
            time.sleep(15 * i)


def modified_today(graph, drive_id: str, settings: Settings) -> bool:
    from sharepoint.client import GraphRequestError

    for name in (f"{REPORTS[0].file_base}.xlsm", f"{REPORTS[0].file_base}.xlsx"):
        try:
            item = graph.get_item_by_path(drive_id, f"{settings.sp_folder}/{name}")
        except GraphRequestError as exc:
            if exc.status_code == 404:
                continue
            raise
        stamp = datetime.fromisoformat(item["lastModifiedDateTime"].replace("Z", "+00:00"))
        return stamp.astimezone(VN_TZ).date() == now_vn().date()
    return False


# ---- một lượt lấy báo cáo -------------------------------------------------------
def run_job(settings: Settings, tu, den, *, bot=None, chat_id=None, dry_run=False) -> int:
    from .fit import FitSession

    def say(text):
        print("[Bot] " + text.splitlines()[0][:120])
        if bot and chat_id:
            try:
                bot.send_message(chat_id, text)
            except FitError as exc:
                print(f"[Telegram] {exc}")

    def captcha(png: bytes, attempt: int):
        if bot and chat_id:
            caption = ("Mã CAPTCHA FIT" + (" (nhập lại, lần trước sai)" if attempt > 1 else "")
                       + f" — trả lời ảnh này bằng mã trong {settings.captcha_wait_s // 60} phút. /huy để dừng.")
            msg_id = bot.send_photo(chat_id, png, caption, force_reply=True)
            print("[Bot] Đã gửi CAPTCHA, đang chờ trả lời…")
            return bot.wait_captcha(chat_id, msg_id, settings.captcha_wait_s)
        out = Path("captcha_fit.png")
        out.write_bytes(png)
        return input(f"Mở {out.resolve()} và nhập mã CAPTCHA: ").strip()

    say(f"Đang lấy ĐH bán trong kỳ NPP {fmt_ky(tu, den)}…")
    shot = None
    try:
        settings.require_fit_credentials()
        graph = drive_id = None
        if not dry_run:
            graph, drive_id = _graph()  # kiểm tra SharePoint trước khi tốn CAPTCHA
        files = []
        with FitSession(settings) as fit:
            try:
                fit.login(captcha)
                fit.open_report()
                for spec in REPORTS:
                    content, suggested, kind = fit.export(spec, tu, den)
                    files.append((spec, target_name(spec, suggested, kind), content))
            except Exception:
                shot = fit.screenshot() if fit.authenticated() else None
                raise
        links = []
        for spec, name, content in files:
            if dry_run:
                Path(name).write_bytes(content)
                links.append(f"{name} (dry-run, lưu tại máy chạy)")
                continue
            item = upload(graph, drive_id, f"{settings.sp_folder}/{name}", content)
            links.append(f"{name}: {item.get('webUrl', '')}")
            print(f"[SharePoint] Đã ghi {settings.sp_folder}/{name} ({len(content):,} byte)")
        if bot and chat_id:
            for spec, name, content in files:
                try:
                    bot.send_document(chat_id, name, content, f"{spec.title} | {fmt_ky(tu, den)}")
                except FitError as exc:
                    print(f"[Telegram] Không gửi được file: {exc}")
        say("✅ Xong " + fmt_ky(tu, den) + "\n" + f"Thư mục SharePoint: {settings.sp_folder}\n" + "\n".join(links))
        return 0
    except FitError as exc:
        say(f"❌ {exc}")
        _send_screen(shot, bot, chat_id)
        return 1
    except Exception as exc:  # noqa: BLE001 - báo loại lỗi, không lộ dữ liệu nhạy cảm
        print(f"[Bot] Lỗi không mong đợi: {type(exc).__name__}: {str(exc)[:300]}")
        say(f"❌ Lỗi {type(exc).__name__} khi lấy báo cáo. Xem log GitHub Actions.")
        _send_screen(shot, bot, chat_id)
        return 1


def _send_screen(png, bot, chat_id):
    """Ảnh màn hình FIT (chỉ khi đã đăng nhập, tức không chứa form mật khẩu) để chẩn đoán từ xa."""
    if png and bot and chat_id:
        try:
            bot.send_photo(chat_id, png, "Màn hình FIT lúc lỗi")
        except FitError:
            pass


# ---- nghe lệnh Telegram ------------------------------------------------------------
def _load_state() -> dict:
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_state(state: dict):
    try:
        STATE_PATH.write_text(json.dumps(state), encoding="utf-8")
    except OSError:
        pass


def _auto_due(settings: Settings, state: dict) -> bool:
    if settings.auto_gio is None:
        return False
    now = now_vn()
    return now.time() >= settings.auto_gio and state.get("auto_date") != now.date().isoformat()


def handle_message(settings, bot, msg) -> None:
    chat_id = msg["chat"]["id"]
    cmd = parse_command(msg["text"])
    if cmd is None:
        bot.send_message(chat_id, "Gõ /banhang để lấy báo cáo, /help xem hướng dẫn.")
        return
    name, arg = cmd
    if name in ("start", "help"):
        bot.send_message(chat_id, HELP)
    elif name in ("banhang", "bh"):
        try:
            tu, den = parse_ky(arg)
        except FitError as exc:
            bot.send_message(chat_id, str(exc))
            return
        run_job(settings, tu, den, bot=bot, chat_id=chat_id)
    elif name in ("huy", "cancel"):
        bot.send_message(chat_id, "Không có lượt nào đang chờ CAPTCHA.")
    else:
        bot.send_message(chat_id, "Lệnh chưa hỗ trợ. /help xem hướng dẫn.")


def listen(settings: Settings, seconds: int) -> int:
    deadline = time.monotonic() + seconds
    if not settings.telegram_enabled:
        print("[Bot] Chưa cấu hình TELEGRAM_BOT_TOKEN / TELEGRAM_ALLOWED_USER_IDS -> chỉ chờ.")
        time.sleep(max(0, seconds))
        return 0
    from .telegram import TelegramBot

    bot = TelegramBot(settings.telegram_token, settings.telegram_user_ids)
    state = _load_state()
    try:
        while True:
            if _auto_due(settings, state):
                state["auto_date"] = now_vn().date().isoformat()
                _save_state(state)
                try:
                    graph, drive_id = _graph()
                    skip = modified_today(graph, drive_id, settings)
                except Exception as exc:  # noqa: BLE001
                    print(f"[Bot] Không kiểm tra được SharePoint ({type(exc).__name__}) -> vẫn chạy.")
                    skip = False
                if skip:
                    print("[Bot] Hôm nay đã có file -> bỏ qua lượt tự động.")
                else:
                    tu, den = parse_ky("")
                    run_job(settings, tu, den, bot=bot, chat_id=min(settings.telegram_user_ids))
            remaining = int(deadline - time.monotonic())
            if remaining <= 0:
                break
            for upd in bot.get_updates(timeout=min(25, remaining)):
                msg = upd.get("message") or {}
                if (msg.get("chat") or {}).get("type") == "private" and (parse_command(msg.get("text", "")) or ("",))[0] == "id":
                    bot.send_message(msg["chat"]["id"], f"ID Telegram của bạn: {msg['from']['id']}")
                    continue
                allowed = bot.authorized_message(upd)
                if allowed is not None:
                    handle_message(settings, bot, allowed)
    finally:
        bot.confirm()
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="python -m fit_ban_hang", description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run", help="Chạy một lượt")
    r.add_argument("--ky", default="", help='"" | thangtruoc | 10/2026 | "01/10/2026 15/10/2026"')
    r.add_argument("--dry-run", action="store_true", help="Không ghi SharePoint, lưu file tại thư mục hiện tại")
    r.add_argument("--no-telegram", action="store_true", help="Nhập CAPTCHA tại terminal")
    li = sub.add_parser("listen", help="Nghe lệnh Telegram trong N giây")
    li.add_argument("--seconds", type=int, default=900)
    args = ap.parse_args(argv)
    try:
        settings = Settings.from_env()
        if args.cmd == "listen":
            return listen(settings, args.seconds)
        tu, den = parse_ky(args.ky)
        bot = chat_id = None
        if settings.telegram_enabled and not args.no_telegram:
            from .telegram import TelegramBot

            bot = TelegramBot(settings.telegram_token, settings.telegram_user_ids)
            chat_id = min(settings.telegram_user_ids)
        try:
            return run_job(settings, tu, den, bot=bot, chat_id=chat_id, dry_run=args.dry_run)
        finally:
            if bot:
                bot.confirm()
    except FitError as exc:
        print(f"[Lỗi] {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
