import calendar
import re
import time
import zipfile
from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo

from lxml import etree

from sync_planning_fc import (
    _find_sheet_xml_path,
    _get_or_create_cell,
    _load_shared_strings,
    _read_cell_text,
)
from sync_stock import DEST_PATH, GraphClient, get_access_token, is_retryable_graph_error

FC_SHEET = "FC"
FC_SELECTOR_CELL = "R1"
PLANNING_SHEET = "Ke_hoach_SX"
START_COLUMN_NUMBER = 19  # S
MAX_DAYS = 31
TIMEZONE = ZoneInfo("Asia/Ho_Chi_Minh")


def _column_letter(number):
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def parse_plan_month(selector):
    text = str(selector or "").strip()
    match = re.search(r"(?:tháng\s*)?(\d{1,2})", text, flags=re.IGNORECASE)
    if not match:
        raise RuntimeError(
            f"{FC_SHEET}!{FC_SELECTOR_CELL}={selector!r} không xác định được tháng."
        )

    month = int(match.group(1))
    if not 1 <= month <= 12:
        raise RuntimeError(
            f"{FC_SHEET}!{FC_SELECTOR_CELL}={selector!r} có tháng không hợp lệ."
        )
    return month


def weekday_label(value):
    # datetime.weekday(): T2=0 ... T7=5, CN=6
    return "CN" if value.weekday() == 6 else f"T{value.weekday() + 2}"


def build_date_headers(plan_year, plan_month):
    days_in_month = calendar.monthrange(plan_year, plan_month)[1]
    headers = []
    for day in range(1, days_in_month + 1):
        value = datetime(plan_year, plan_month, day)
        headers.append(f"{day:02d}/{plan_month:02d}\n{weekday_label(value)}")
    return headers


def _read_selector_from_archive(archive):
    sheet_path = _find_sheet_xml_path(archive, FC_SHEET)
    shared_strings = _load_shared_strings(archive)
    root = etree.fromstring(archive.read(sheet_path))

    cells = root.xpath(
        f'//*[local-name()="c"][@r="{FC_SELECTOR_CELL}"]'
    )
    if not cells:
        raise RuntimeError(f"Không tìm thấy {FC_SHEET}!{FC_SELECTOR_CELL}.")

    selector = _read_cell_text(cells[0], shared_strings)
    if not selector:
        raise RuntimeError(f"{FC_SHEET}!{FC_SELECTOR_CELL} đang trống.")
    return selector


def _set_inline_text(cell, value):
    for child in list(cell):
        cell.remove(child)

    if value in (None, ""):
        cell.attrib.pop("t", None)
        return

    cell.set("t", "inlineStr")
    namespace = etree.QName(cell).namespace
    inline = etree.SubElement(cell, f"{{{namespace}}}is")
    text_node = etree.SubElement(inline, f"{{{namespace}}}t")
    text_node.set(
        "{http://www.w3.org/XML/1998/namespace}space",
        "preserve",
    )
    text_node.text = value


def prepare_calendar_update(workbook_bytes, *, plan_year=None):
    source_buffer = BytesIO(workbook_bytes)
    output_buffer = BytesIO()

    with zipfile.ZipFile(source_buffer, "r") as source_zip:
        selector = _read_selector_from_archive(source_zip)
        plan_month = parse_plan_month(selector)
        if plan_year is None:
            plan_year = datetime.now(TIMEZONE).year

        headers = build_date_headers(plan_year, plan_month)
        targets = headers + [""] * (MAX_DAYS - len(headers))

        sheet_path = _find_sheet_xml_path(source_zip, PLANNING_SHEET)
        shared_strings = _load_shared_strings(source_zip)
        sheet_root = etree.fromstring(source_zip.read(sheet_path))

        row_nodes = sheet_root.xpath(
            '//*[local-name()="sheetData"]/*[local-name()="row"][@r="1"]'
        )
        if not row_nodes:
            raise RuntimeError(f"Không tìm thấy dòng 1 của {PLANNING_SHEET}.")
        row = row_nodes[0]

        # Dùng style của R1 cho toàn bộ dải ngày để đồng bộ hàng tiêu đề.
        r1_nodes = row.xpath('./*[local-name()="c"][@r="R1"]')
        header_style = r1_nodes[0].get("s") if r1_nodes else None

        changed = 0
        for index, target in enumerate(targets):
            column_number = START_COLUMN_NUMBER + index
            column_letter = _column_letter(column_number)
            ref = f"{column_letter}1"

            existing = row.xpath(f'./*[local-name()="c"][@r="{ref}"]')
            current = (
                _read_cell_text(existing[0], shared_strings)
                if existing
                else ""
            )

            if current == target:
                # Vẫn chuẩn hóa style nếu cell đã tồn tại.
                if existing and header_style is not None:
                    existing[0].set("s", header_style)
                continue

            cell = existing[0] if existing else _get_or_create_cell(
                row,
                1,
                column_letter,
            )
            if header_style is not None:
                cell.set("s", header_style)
            _set_inline_text(cell, target)
            changed += 1

        if changed == 0:
            return workbook_bytes, {
                "selector": selector,
                "plan_year": plan_year,
                "plan_month": plan_month,
                "days": len(headers),
                "start": "S1",
                "end": f"{_column_letter(START_COLUMN_NUMBER + len(headers) - 1)}1",
                "changed_count": 0,
            }

        new_sheet_xml = etree.tostring(
            sheet_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        with zipfile.ZipFile(
            output_buffer,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as output_zip:
            for item in source_zip.infolist():
                data = (
                    new_sheet_xml
                    if item.filename == sheet_path
                    else source_zip.read(item.filename)
                )
                output_zip.writestr(item, data)

    return output_buffer.getvalue(), {
        "selector": selector,
        "plan_year": plan_year,
        "plan_month": plan_month,
        "days": len(headers),
        "start": "S1",
        "end": f"{_column_letter(START_COLUMN_NUMBER + len(headers) - 1)}1",
        "changed_count": changed,
    }


def main_with_retry(
    *,
    sleep_func=time.sleep,
    max_attempts=6,
    retry_delay_seconds=10,
):
    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    for attempt in range(1, max_attempts + 1):
        dest_item = graph.get_item_by_path(drive_id, DEST_PATH)
        dest_bytes = graph.download_file(drive_id, dest_item["id"])
        updated_bytes, info = prepare_calendar_update(dest_bytes)

        print(
            f"[{PLANNING_SHEET}] {FC_SHEET}!{FC_SELECTOR_CELL}="
            f"{info['selector']!r} -> {info['plan_month']:02d}/{info['plan_year']}; "
            f"dải ngày {info['start']}:{info['end']} ({info['days']} ngày)."
        )

        if info["changed_count"] == 0:
            print(f"[{PLANNING_SHEET}] Dải ngày đã đúng; không cần upload lại.")
            return

        try:
            graph.upload_file(
                drive_id,
                dest_item["id"],
                updated_bytes,
                expected_etag=dest_item["eTag"],
            )
            print(
                f"[{PLANNING_SHEET}] Đã cập nhật {info['changed_count']} ô tiêu đề ngày."
            )
            return
        except Exception as exc:
            if not is_retryable_graph_error(exc) or attempt == max_attempts:
                raise

            print(
                f"[{PLANNING_SHEET}] File đang khóa/thay đổi; "
                f"thử lại ({attempt}/{max_attempts})."
            )
            sleep_func(retry_delay_seconds)


if __name__ == "__main__":
    main_with_retry()
