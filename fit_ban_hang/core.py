"""Phần thuần (không mạng): cấu hình, kỳ báo cáo, kiểm tra/đặt tên file Excel."""

from __future__ import annotations

import calendar
import io
import os
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from urllib.parse import urlsplit

VN_TZ = timezone(timedelta(hours=7))
MAX_EXCEL_BYTES = 45 * 1024 * 1024  # Telegram cho bot gửi tối đa 50 MB


class FitError(RuntimeError):
    """Lỗi có thể báo thẳng cho người dùng (không chứa token/mật khẩu)."""


@dataclass(frozen=True)
class ReportSpec:
    key: str          # tham số `muc` của API Worker
    muc_lay: str      # nhãn "Mức lấy" trên FIT
    file_base: str    # tên file trên SharePoint (không gồm đuôi)
    title: str


# Ảnh 1: Mức lấy = Đơn hàng -> BCDonHangBanTrongKyNPP ; Ảnh 2: Mức lấy = Hóa Đơn -> Lay So Hoa don
REPORTS = (
    ReportSpec("don_hang", "Đơn hàng", "BCDonHangBanTrongKyNPP", "ĐH bán trong kỳ NPP – theo Đơn hàng"),
    ReportSpec("hoa_don", "Hóa Đơn", "Lay So Hoa don", "ĐH bán trong kỳ NPP – theo Hóa đơn"),
)

WORKER_ORIGIN = "https://erp-nhan-hang.nguyensihuan711994.workers.dev"
WORKER_PATH = "/api/integrations/ban-hang-npp/file"
DEFAULT_SP_FOLDER = "Planning App/Bao cao ban hang/DH Ban trong ky NPP"


def _env(name: str, default: str = "") -> str:
    return str(os.environ.get(name, default) or default).strip()


@dataclass(frozen=True)
class Settings:
    cf_client_id: str = field(default="", repr=False)
    cf_client_secret: str = field(default="", repr=False)
    worker_origin: str = WORKER_ORIGIN
    sp_folder: str = DEFAULT_SP_FOLDER
    telegram_token: str = field(default="", repr=False)
    telegram_chat_id: str = ""

    @classmethod
    def from_env(cls) -> "Settings":
        origin = _env("FIT_WORKER_ORIGIN", WORKER_ORIGIN).rstrip("/")
        parts = urlsplit(origin)
        if parts.scheme != "https" or not parts.hostname or parts.path or parts.query:
            raise FitError("FIT_WORKER_ORIGIN phải là https://<host>, không kèm đường dẫn.")
        chat = _env("TELEGRAM_CHAT_ID")
        if chat and not re.fullmatch(r"-?\d{3,20}", chat):
            raise FitError("TELEGRAM_CHAT_ID phải là số.")
        return cls(
            cf_client_id=_env("FIT_CF_ACCESS_CLIENT_ID"),
            cf_client_secret=_env("FIT_CF_ACCESS_CLIENT_SECRET"),
            worker_origin=origin,
            sp_folder=_env("FIT_SP_FOLDER", DEFAULT_SP_FOLDER).strip("/"),
            telegram_token=_env("TELEGRAM_BOT_TOKEN"),
            telegram_chat_id=chat,
        )

    def require_worker_credentials(self):
        missing = [k for k, v in (("FIT_CF_ACCESS_CLIENT_ID", self.cf_client_id),
                                  ("FIT_CF_ACCESS_CLIENT_SECRET", self.cf_client_secret)) if not v]
        if missing:
            raise FitError("Thiếu secret GitHub: " + ", ".join(missing))

    @property
    def telegram_enabled(self) -> bool:
        return bool(self.telegram_token and self.telegram_chat_id)


def now_vn() -> datetime:
    return datetime.now(VN_TZ)


def _parse_date(text: str) -> date:
    text = text.strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise FitError(f"Ngày không hợp lệ: {text!r}. Dùng dd/mm/yyyy.")


def _month_range(year: int, month: int, today: date) -> tuple[date, date]:
    if not 1 <= month <= 12:
        raise FitError("Tháng phải từ 1 đến 12.")
    first = date(year, month, 1)
    last = date(year, month, calendar.monthrange(year, month)[1])
    if first > today:
        raise FitError("Tháng này chưa bắt đầu.")
    return first, min(last, today)


def parse_ky(text: str, today: date | None = None) -> tuple[date, date]:
    """Kỳ báo cáo (giống lệnh Telegram /banhang của Worker).

    ""                       -> đầu tháng đến hôm nay
    "thangtruoc"             -> cả tháng trước
    "10/2026" | "2026-10"    -> cả tháng (tháng hiện tại thì đến hôm nay)
    "01/10/2026"             -> một ngày
    "01/10/2026 15/10/2026"  -> khoảng ngày (cũng nhận "2026-10-01 2026-10-15")
    """
    today = today or now_vn().date()
    raw = " ".join((text or "").split())
    low = raw.lower()
    if not raw:
        return today.replace(day=1), today
    if low in ("thangtruoc", "thang-truoc", "tháng trước", "thang truoc", "tt"):
        prev = today.replace(day=1) - timedelta(days=1)
        return _month_range(prev.year, prev.month, today)
    m = re.fullmatch(r"(\d{1,2})[/-](\d{4})", raw)
    if m:
        return _month_range(int(m[2]), int(m[1]), today)
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", raw)
    if m:
        return _month_range(int(m[1]), int(m[2]), today)
    tokens = re.sub(r"\s*(?:->|→|–|—)\s*|\s+(?:-|den|đến)\s+", " ", raw).split()
    if len(tokens) == 1:
        tu = den = _parse_date(tokens[0])
    elif len(tokens) == 2:
        tu, den = _parse_date(tokens[0]), _parse_date(tokens[1])
    else:
        raise FitError("Kỳ không hợp lệ. Ví dụ: 10/2026 | thangtruoc | 01/10/2026 15/10/2026")
    if tu > den:
        raise FitError("Từ ngày phải trước hoặc bằng Đến ngày.")
    if den > today:
        raise FitError("Đến ngày không được sau hôm nay.")
    if (den - tu).days > 366:
        raise FitError("Kỳ báo cáo tối đa 1 năm.")
    return tu, den


def detect_excel_kind(content: bytes) -> str | None:
    """'ooxml' (xlsx/xlsm), 'xls' (BIFF/SpreadsheetML) hoặc None nếu là trang HTML/lỗi."""
    if not content:
        return None
    if content[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                if {"[Content_Types].xml", "xl/workbook.xml"}.issubset(zf.namelist()):
                    return "ooxml"
        except zipfile.BadZipFile:
            return None
        return None
    if content[:8] == bytes.fromhex("D0CF11E0A1B11AE1"):
        return "xls"
    if b"urn:schemas-microsoft-com:office:spreadsheet" in content[:4096]:
        return "xls"
    return None


def validate_excel(content: bytes, label: str) -> str:
    if not content:
        raise FitError(f"{label}: file tải về rỗng.")
    if len(content) > MAX_EXCEL_BYTES:
        raise FitError(f"{label}: file vượt 45 MB.")
    kind = detect_excel_kind(content)
    if kind is None:
        raise FitError(f"{label}: FIT trả về nội dung không phải Excel.")
    return kind


def ooxml_extension(content: bytes) -> str:
    """'.xlsm' nếu workbook là loại macro-enabled (FIT xuất .xlsm), ngược lại '.xlsx'."""
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as zf:
            types = zf.read("[Content_Types].xml")
    except (zipfile.BadZipFile, KeyError):
        return ".xlsx"
    return ".xlsm" if b"macroEnabled" in types else ".xlsx"


def target_name(spec: ReportSpec, content: bytes, kind: str) -> str:
    """Tên lưu SharePoint đúng ảnh 3; đuôi theo nội dung thật để Excel mở được."""
    ext = ooxml_extension(content) if kind == "ooxml" else ".xls"
    return spec.file_base + ext


def fmt_ky(tu: date, den: date) -> str:
    return f"{tu:%d/%m/%Y} – {den:%d/%m/%Y}"
