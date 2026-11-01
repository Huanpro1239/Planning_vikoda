"""FIT "ĐH Bán trong kỳ NPP" -> SharePoint (qua Worker erp-nhan-hang), báo kết quả vào Telegram.

    python -X utf8 -m fit_ban_hang                         # đầu tháng -> hôm nay
    python -X utf8 -m fit_ban_hang --ky thangtruoc
    python -X utf8 -m fit_ban_hang --start 2026-10-01 --end 2026-10-31
    python -X utf8 -m fit_ban_hang --dry-run               # không ghi SharePoint, lưu file tại chỗ

Kích hoạt: lệnh Telegram /banhang (Worker gửi repository_dispatch fit_ban_hang) hoặc chạy tay
workflow "FIT Ban Hang NPP". Ghi đè BCDonHangBanTrongKyNPP.xlsm (Mức lấy Đơn hàng) và
Lay So Hoa don.xlsm (Mức lấy Hóa Đơn) trong FIT_SP_FOLDER.
"""

from __future__ import annotations

import argparse
import sys
import time
from datetime import date
from pathlib import Path
from urllib.parse import quote

from .core import REPORTS, FitError, Settings, fmt_ky, parse_ky, target_name

XLSM_CT = "application/vnd.ms-excel.sheet.macroEnabled.12"


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


def run(settings: Settings, tu: date, den: date, *, dry_run=False, client=None, notifier=None,
        graph_factory=_graph) -> int:
    from .worker import WorkerClient

    if notifier is None and settings.telegram_enabled:
        from .telegram import TelegramNotifier

        notifier = TelegramNotifier(settings.telegram_token, settings.telegram_chat_id)

    def say(text):
        print("[Bot] " + text.splitlines()[0][:150])
        if notifier:
            try:
                notifier.send_message(text)
            except FitError as exc:
                print(f"[Telegram] {exc}")

    try:
        client = client or WorkerClient(settings)
        files = []
        for spec in REPORTS:
            content, kind = client.fetch(spec, tu, den)
            files.append((spec, target_name(spec, content, kind), content))
        lines = []
        if dry_run:
            for spec, name, content in files:
                Path(name).write_bytes(content)
                lines.append(f"{name} (dry-run, lưu tại máy chạy)")
        else:
            graph, drive_id = graph_factory()
            for spec, name, content in files:
                item = upload(graph, drive_id, f"{settings.sp_folder}/{name}", content)
                print(f"[SharePoint] Đã ghi {settings.sp_folder}/{name} ({len(content):,} byte)")
                lines.append(f"{name}: {item.get('webUrl', '')}")
        if notifier:
            for spec, name, content in files:
                try:
                    notifier.send_document(name, content, f"{spec.title} | {fmt_ky(tu, den)}")
                except FitError as exc:
                    print(f"[Telegram] Không gửi được file: {exc}")
        say(f"✅ ĐH bán trong kỳ NPP {fmt_ky(tu, den)} đã lên SharePoint\n"
            f"Thư mục: {settings.sp_folder}\n" + "\n".join(lines))
        return 0
    except FitError as exc:
        say(f"❌ ĐH bán trong kỳ NPP {fmt_ky(tu, den)}: {exc}")
        return 1
    except Exception as exc:  # noqa: BLE001 - chỉ báo loại lỗi
        print(f"[Bot] Lỗi không mong đợi: {type(exc).__name__}: {str(exc)[:300]}")
        say(f"❌ ĐH bán trong kỳ NPP {fmt_ky(tu, den)}: lỗi {type(exc).__name__}. Xem log GitHub Actions.")
        return 1


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="python -m fit_ban_hang", description=__doc__.splitlines()[0])
    ap.add_argument("--ky", default="", help='"" | thangtruoc | 10/2026 | "01/10/2026 15/10/2026"')
    ap.add_argument("--start", default="", help="YYYY-MM-DD (dùng cùng --end, thay cho --ky)")
    ap.add_argument("--end", default="")
    ap.add_argument("--dry-run", action="store_true", help="Không ghi SharePoint")
    args = ap.parse_args(argv)
    try:
        settings = Settings.from_env()
        if args.start or args.end:
            tu, den = parse_ky(f"{args.start} {args.end or args.start}".strip())
        else:
            tu, den = parse_ky(args.ky)
        return run(settings, tu, den, dry_run=args.dry_run)
    except FitError as exc:
        print(f"[Lỗi] {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
