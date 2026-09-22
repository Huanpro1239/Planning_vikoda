"""OpenXML patch/compare helpers for Planning metrics."""

import math
import zipfile
from datetime import date, datetime
from io import BytesIO

from lxml import etree

from sync_planning_fc import (
    _find_sheet_xml_path,
    _get_or_create_cell,
    _load_shared_strings,
    _read_cell_text,
)
from stock import clean_number, normalize_code

from .constants import OUTPUT_COLUMNS, PLANNING_SHEET


def datetime_to_excel_serial(value):
    if value is None:
        return None
    epoch = datetime(1899, 12, 30)
    return (value - epoch).total_seconds() / 86400


def set_cell_value(row_element, row_number, column_letter, value):
    cell = _get_or_create_cell(row_element, row_number, column_letter)
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)

    if value is None:
        return
    if isinstance(value, datetime):
        value = datetime_to_excel_serial(value)

    namespace = etree.QName(cell).namespace
    value_node = etree.SubElement(cell, f"{{{namespace}}}v")
    value = clean_number(value)
    value_node.text = (
        str(value)
        if isinstance(value, int)
        else format(float(value), ".15g")
    )


def patch_workbook(workbook_bytes, metric_values):
    source_buffer = BytesIO(workbook_bytes)
    output_buffer = BytesIO()

    with zipfile.ZipFile(source_buffer, "r") as source_zip:
        sheet_path = _find_sheet_xml_path(source_zip, PLANNING_SHEET)
        shared_strings = _load_shared_strings(source_zip)
        sheet_root = etree.fromstring(source_zip.read(sheet_path))

        seen = set()
        patched = 0
        rows = sheet_root.xpath(
            '//*[local-name()="sheetData"]/*[local-name()="row"]'
        )

        for row_element in rows:
            row_number_text = row_element.get("r")
            if not row_number_text:
                continue
            row_number = int(row_number_text)

            code = None
            for cell in row_element.xpath('./*[local-name()="c"]'):
                if cell.get("r") == f"A{row_number}":
                    code = normalize_code(
                        _read_cell_text(cell, shared_strings)
                    )
                    break

            if not code or code not in metric_values:
                continue
            if code in seen:
                raise RuntimeError(
                    f"Mã {code} bị lặp trong XML {PLANNING_SHEET}!A."
                )
            seen.add(code)

            for column, key in OUTPUT_COLUMNS.items():
                set_cell_value(
                    row_element,
                    row_number,
                    column,
                    metric_values[code][key],
                )
            patched += 1

        if patched != len(metric_values):
            missing = sorted(set(metric_values) - seen)
            raise RuntimeError(
                f"Chỉ cập nhật {patched}/{len(metric_values)} mã vào "
                f"{PLANNING_SHEET}!M:R. Thiếu: "
                + ", ".join(missing[:10])
            )

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

    return output_buffer.getvalue(), patched


def same_value(current, target):
    if target is None:
        return current in (None, "")

    if isinstance(target, datetime):
        if isinstance(current, datetime):
            return abs((current - target).total_seconds()) < 1
        if isinstance(current, date):
            return current == target.date()
        return False

    if isinstance(current, (int, float)) and isinstance(target, (int, float)):
        return math.isclose(
            float(current),
            float(target),
            rel_tol=1e-10,
            abs_tol=1e-7,
        )
    return current == target


def count_changes(planning_rows, metric_values):
    changed = 0
    for code, row in planning_rows.items():
        current = row["current_outputs"]
        target = metric_values[code]
        if any(
            not same_value(current[key], target[key])
            for key in OUTPUT_COLUMNS.values()
        ):
            changed += 1
    return changed
