import re
import zipfile
from io import BytesIO

from lxml import etree

from sync_planning_fc import _find_sheet_xml_path, _load_shared_strings, _read_cell_text


PLANNING_SHEET = "Ke_hoach_SX"
START_COLUMN_NUMBER = 19  # S


def _column_number_from_ref(ref):
    match = re.match(r"([A-Z]+)", str(ref or "").upper())
    if not match:
        return None
    value = 0
    for char in match.group(1):
        value = value * 26 + ord(char) - 64
    return value


def _column_letter(number):
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _looks_like_date_header(value):
    text = str(value or "").strip()
    first_line = text.splitlines()[0] if text else ""
    return bool(re.fullmatch(r"\d{2}/\d{2}", first_line))


def needs_schedule_reset(workbook_bytes):
    """Legacy layout had S='Kỳ kế hoạch'; canonical layout has a date in S1."""
    with zipfile.ZipFile(BytesIO(workbook_bytes), "r") as archive:
        sheet_path = _find_sheet_xml_path(archive, PLANNING_SHEET)
        shared_strings = _load_shared_strings(archive)
        root = etree.fromstring(archive.read(sheet_path))
        nodes = root.xpath('//*[local-name()="c"][@r="S1"]')
        if not nodes:
            return True
        return not _looks_like_date_header(_read_cell_text(nodes[0], shared_strings))


def _normalize_cols(root, active_end):
    cols_nodes = root.xpath('//*[local-name()="cols"]')
    if not cols_nodes:
        return 0

    cols = cols_nodes[0]
    children = [child for child in list(cols) if etree.QName(child).localname == "col"]
    schedule_cols = []
    for col in children:
        minimum = int(col.get("min", "1"))
        maximum = int(col.get("max", str(minimum)))
        if maximum >= START_COLUMN_NUMBER:
            schedule_cols.append(col)

    # Already canonical: exactly one schedule-format definition S:last-day,
    # and no formatting survives past the active date range.
    if len(schedule_cols) == 1:
        only = schedule_cols[0]
        if (
            int(only.get("min", "0")) == START_COLUMN_NUMBER
            and int(only.get("max", "0")) == active_end
        ):
            return 0

    date_attrs = None
    kept = []
    for col in children:
        minimum = int(col.get("min", "1"))
        maximum = int(col.get("max", str(minimum)))
        if maximum < START_COLUMN_NUMBER:
            kept.append(col)
            continue

        if date_attrs is None and minimum <= 20 <= maximum:
            date_attrs = dict(col.attrib)
        elif date_attrs is None and minimum <= START_COLUMN_NUMBER <= maximum:
            date_attrs = dict(col.attrib)

        if minimum < START_COLUMN_NUMBER:
            col.set("max", str(START_COLUMN_NUMBER - 1))
            kept.append(col)

    if date_attrs is None:
        date_attrs = {"width": "8.88671875", "customWidth": "1"}
    date_attrs.pop("min", None)
    date_attrs.pop("max", None)

    namespace = etree.QName(cols).namespace
    date_col = etree.Element(f"{{{namespace}}}col")
    date_col.set("min", str(START_COLUMN_NUMBER))
    date_col.set("max", str(active_end))
    for key, value in date_attrs.items():
        date_col.set(key, value)
    kept.append(date_col)

    for child in list(cols):
        if etree.QName(child).localname == "col":
            cols.remove(child)
    kept.sort(key=lambda node: int(node.get("min", "0")))
    for child in kept:
        cols.append(child)
    return 1


def canonicalize_planning_layout(workbook_bytes, *, active_days, reset_schedule=False):
    """
    Keep A:R plus exactly the active date range S:...

    Removes legacy 'Kỳ kế hoạch', fixed Ngày 31, Tổng SX, Chênh lệch and
    formatting/cells beyond the last actual day of the selected month.
    """
    if not 1 <= int(active_days) <= 31:
        raise ValueError("active_days phải từ 1 đến 31.")

    active_end = START_COLUMN_NUMBER + int(active_days) - 1
    active_end_letter = _column_letter(active_end)
    source_buffer = BytesIO(workbook_bytes)
    output_buffer = BytesIO()

    with zipfile.ZipFile(source_buffer, "r") as source_zip:
        sheet_path = _find_sheet_xml_path(source_zip, PLANNING_SHEET)
        root = etree.fromstring(source_zip.read(sheet_path))
        changed = 0

        sheet_data_nodes = root.xpath('//*[local-name()="sheetData"]')
        if not sheet_data_nodes:
            raise RuntimeError(f"Không tìm thấy sheetData của {PLANNING_SHEET}.")
        row_nodes = sheet_data_nodes[0].xpath('./*[local-name()="row"]')
        last_row = 1

        for row in row_nodes:
            row_number = int(row.get("r", "1"))
            last_row = max(last_row, row_number)
            for cell in list(row.xpath('./*[local-name()="c"]')):
                col_num = _column_number_from_ref(cell.get("r"))
                if col_num is None:
                    continue
                remove = col_num > active_end
                if reset_schedule and row_number > 1 and col_num >= START_COLUMN_NUMBER:
                    remove = True
                if remove:
                    row.remove(cell)
                    changed += 1

            desired_span = f"1:{active_end}"
            if row.get("spans") is not None and row.get("spans") != desired_span:
                row.set("spans", desired_span)
                changed += 1

        desired_ref = f"A1:{active_end_letter}{last_row}"
        dimension_nodes = root.xpath('//*[local-name()="dimension"]')
        if dimension_nodes and dimension_nodes[0].get("ref") != desired_ref:
            dimension_nodes[0].set("ref", desired_ref)
            changed += 1

        filter_nodes = root.xpath('//*[local-name()="autoFilter"]')
        if filter_nodes and filter_nodes[0].get("ref") != desired_ref:
            filter_nodes[0].set("ref", desired_ref)
            changed += 1

        changed += _normalize_cols(root, active_end)

        if changed == 0:
            return workbook_bytes, 0

        new_xml = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        with zipfile.ZipFile(output_buffer, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
            for item in source_zip.infolist():
                data = new_xml if item.filename == sheet_path else source_zip.read(item.filename)
                output_zip.writestr(item, data)

    return output_buffer.getvalue(), changed
