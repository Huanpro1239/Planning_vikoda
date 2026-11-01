"""Phần thuần (không mạng): cấu hình, kỳ báo cáo, lệnh Telegram, kiểm tra file Excel."""

from __future__ import annotations

import calendar
import io
import os
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, time as dtime, timedelta, timezone
from urllib.parse import urljoin, urlsplit

VN_TZ = timezone(timedelta(hours=7))
MAX_EXCEL_BYTES = 45 * 1024 * 1024  # Telegram cho bot gửi tối đa 50 MB


class FitError(RuntimeError):
    """Lỗi có thể báo thẳng cho người dùng (không chứa tài khoản/mật khẩu)."""


@dataclass(frozen=True)
class ReportSpec:
    key: str
    muc_lay: str      # nhãn radio "Mức lấy" trên FIT
    file_base: str    # tên file trên SharePoint (không gồm đuôi)
    title: str


# Ảnh 1: Mức lấy = Đơn hàng -> BCDonHangBanTrongKyNPP ; Ảnh 2: Mức lấy = Hóa Đơn -> Lay So Hoa don
REPORTS = (
    ReportSpec("don_hang", "Đơn hàng", "BCDonHangBanTrongKyNPP", "ĐH bán trong kỳ NPP – theo Đơn hàng"),
    ReportSpec("hoa_don", "Hóa Đơn", "Lay So Hoa don", "ĐH bán trong kỳ NPP – theo Hóa đơn"),
)

MENU_PATH = ("Báo cáo kế toán", "Báo cáo khác", "ĐH Bán trong kỳ NPP")
DEFAULT_BASE_URL = "https://103.68.251.239:8888/FIT/"
DEFAULT_SP_FOLDER = "Planning App/Bao cao ban hang/DH Ban trong ky NPP"


def _env(name: str, default: str = "") -> str:
    return str(os.environ.get(name, default) or default).strip()


def _int_env(name: str, default: int) -> int:
    raw = _env(name, str(default))
    try:
        value = int(raw)
    except ValueError:
        raise FitError(f"{name} phải là số nguyên.") from None
    if value <= 0:
        raise FitError(f"{name} phải lớn hơn 0.")
    return value


def parse_user_ids(raw: str) -> frozenset[int]:
    ids = set()
    for part in re.split(r"[,\s;]+", raw or ""):
        if not part:
            continue
        if not re.fullmatch(r"\d{3,15}", part):
            raise FitError("TELEGRAM_ALLOWED_USER_IDS chỉ gồm ID số, ngăn cách bằng dấu phẩy.")
        ids.add(int(part))
    return frozenset(ids)


def parse_gio(raw: str) -> dtime | None:
    raw = (raw or "").strip()
    if not raw:
        return None
    m = re.fullmatch(r"(\d{1,2})[:hH](\d{2})", raw)
    if not m or int(m[1]) > 23 or int(m[2]) > 59:
        raise FitError("FIT_AUTO_GIO phải dạng HH:MM, ví dụ 07:30.")
    return dtime(int(m[1]), int(m[2]))


@dataclass(frozen=True)
class Settings:
    base_url: str = DEFAULT_BASE_URL
    username: str = field(default="", repr=False)
    password: str = field(default="", repr=False)
    sp_folder: str = DEFAULT_SP_FOLDER
    report_url: str = ""
    date_format: str = "%d/%m/%Y"
    ignore_https_errors: bool = False
    headless: bool = True
    timeout_ms: int = 45000
    download_timeout_ms: int = 240000
    captcha_wait_s: int = 300
    auto_gio: dtime | None = None
    telegram_token: str = field(default="", repr=False)
    telegram_user_ids: frozenset[int] = frozenset()

    @classmethod
    def from_env(cls) -> "Settings":
        base = _env("FIT_BASE_URL", DEFAULT_BASE_URL).rstrip("/") + "/"
        parts = urlsplit(base)
        if parts.scheme not in ("http", "https") or not parts.hostname or parts.username or parts.password:
            raise FitError("FIT_BASE_URL phải là URL http(s) không chứa tài khoản.")
        flag = _env("FIT_IGNORE_HTTPS_ERRORS", "false").lower()
        headless = _env("FIT_HEADLESS", "true").lower()
        if flag not in ("true", "false") or headless not in ("true", "false"):
            raise FitError("FIT_IGNORE_HTTPS_ERRORS / FIT_HEADLESS phải là true hoặc false.")
        return cls(
            base_url=base,
            username=_env("FIT_USERNAME"),
            password=os.environ.get("FIT_PASSWORD", ""),
            sp_folder=_env("FIT_SP_FOLDER", DEFAULT_SP_FOLDER).strip("/"),
            report_url=_env("FIT_REPORT_URL"),
            date_format=_env("FIT_DATE_FORMAT", "%d/%m/%Y"),
            ignore_https_errors=flag == "true",
            headless=headless == "true",
            timeout_ms=_int_env("FIT_TIMEOUT_MS", 45000),
            download_timeout_ms=_int_env("FIT_DOWNLOAD_TIMEOUT_MS", 240000),
            captcha_wait_s=_int_env("FIT_CAPTCHA_WAIT", 300),
            auto_gio=parse_gio(_env("FIT_AUTO_GIO")),
            telegram_token=_env("TELEGRAM_BOT_TOKEN"),
            telegram_user_ids=parse_user_ids(_env("TELEGRAM_ALLOWED_USER_IDS")),
        )

    def require_fit_credentials(self):
        missing = [k for k, v in (("FIT_USERNAME", self.username), ("FIT_PASSWORD", self.password)) if not v]
        if missing:
            raise FitError("Thiếu secret: " + ", ".join(missing))

    @property
    def telegram_enabled(self) -> bool:
        return bool(self.telegram_token and self.telegram_user_ids)


def same_site_url(base_url: str, target: str) -> str:
    url = urljoin(base_url, target)
    a, b = urlsplit(base_url), urlsplit(url)
    if (a.scheme, a.hostname, a.port) != (b.scheme, b.hostname, b.port) or b.username or b.password:
        raise FitError("URL báo cáo phải cùng máy chủ FIT.")
    return url


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
    """Kỳ báo cáo từ lệnh người dùng.

    ""                       -> đầu tháng đến hôm nay
    "thangtruoc"             -> cả tháng trước
    "10/2026" | "2026-10"    -> cả tháng (tháng hiện tại thì đến hôm nay)
    "01/10/2026"             -> một ngày
    "01/10/2026 15/10/2026"  -> khoảng ngày
    """
    today = today or now_vn().date()
    raw = " ".join((text or "").replace("–", " ").replace("—", " ").split())
    low = raw.lower()
    if not raw:
        return today.replace(day=1), today
    if low in ("thangtruoc", "thang-truoc", "tháng trước", "thang truoc", "tt"):
        prev = today.replace(day=1) - timedelta(days=1)
        return _month_range(prev.year, prev.month, today)
    m = re.fullmatch(r"(\d{1,2})[/-](\d{4})", raw) or None
    if m:
        return _month_range(int(m[2]), int(m[1]), today)
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", raw)
    if m:
        return _month_range(int(m[1]), int(m[2]), today)
    tokens = re.sub(r"\s*(?:->|→)\s*|\s+(?:-|den|đến|toi|tới)\s+", " ", raw).split()
    if len(tokens) == 1:
        d = _parse_date(tokens[0])
        tu, den = d, d
    elif len(tokens) == 2:
        tu, den = _parse_date(tokens[0]), _parse_date(tokens[1])
    else:
        raise FitError("Kỳ không hợp lệ. Ví dụ: /banhang  |  /banhang 10/2026  |  /banhang 01/10/2026 15/10/2026")
    if tu > den:
        raise FitError("Từ ngày phải trước hoặc bằng Đến ngày.")
    if den > today:
        raise FitError("Đến ngày không được sau hôm nay.")
    if (den - tu).days > 366:
        raise FitError("Kỳ báo cáo tối đa 1 năm.")
    return tu, den


def parse_command(text: str) -> tuple[str, str] | None:
    """'/banhang@Bot 10/2026' -> ('banhang', '10/2026'); không phải lệnh -> None."""
    m = re.fullmatch(r"/([A-Za-z_]+)(?:@\w+)?(?:\s+(.*))?", (text or "").strip(), flags=re.S)
    if not m:
        return None
    return m[1].lower(), (m[2] or "").strip()


def looks_like_captcha(text: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9]{3,12}", (text or "").strip()))


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
        raise FitError(f"{label}: FIT trả về nội dung không phải Excel (có thể là trang lỗi/hết phiên).")
    return kind


def target_name(spec: ReportSpec, suggested_filename: str, kind: str) -> str:
    """Tên lưu SharePoint: giữ đúng tên ảnh 3; đuôi theo file thật để Excel mở được."""
    ext = os.path.splitext(suggested_filename or "")[1].lower()
    if kind == "ooxml":
        ext = ext if ext in (".xlsm", ".xlsx") else ".xlsm"
    else:
        ext = ".xls"
    return spec.file_base + ext


def fmt_ky(tu: date, den: date) -> str:
    return f"{tu:%d/%m/%Y} – {den:%d/%m/%Y}"
