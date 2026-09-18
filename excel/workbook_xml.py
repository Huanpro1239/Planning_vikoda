"""Low-level OpenXML helpers shared by Excel patchers.

This module owns workbook ZIP/XML primitives so domain code does not depend on
the legacy sync_stock module.
"""

from __future__ import annotations

import posixpath
import re

from lxml import etree


def load_shared_strings(archive):
    path = "xl/sharedStrings.xml"
    if path not in archive.namelist():
        return []

    root = etree.fromstring(archive.read(path))
    return ["".join(si.itertext()) for si in root.xpath('//*[local-name()="si"]')]


def read_cell_text(cell, shared_strings):
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(cell.itertext()).strip()

    value_nodes = cell.xpath('./*[local-name()="v"]')
    if not value_nodes:
        return ""

    raw = value_nodes[0].text or ""
    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except Exception:
            return ""
    return raw


def find_sheet_xml_path(archive, sheet_name):
    workbook_root = etree.fromstring(archive.read("xl/workbook.xml"))
    relationship_ns = (
        "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships"
    )

    relationship_id = None
    for sheet in workbook_root.xpath('//*[local-name()="sheet"]'):
        if sheet.get("name") == sheet_name:
            relationship_id = sheet.get(f"{{{relationship_ns}}}id")
            break

    if not relationship_id:
        raise RuntimeError(f"Không tìm thấy sheet {sheet_name!r} trong file đích.")

    rels_root = etree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = None
    for rel in rels_root.xpath('//*[local-name()="Relationship"]'):
        if rel.get("Id") == relationship_id:
            target = rel.get("Target")
            break

    if not target:
        raise RuntimeError(f"Không xác định được XML của sheet {sheet_name!r}.")

    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


def column_number(cell_reference):
    letters = re.match(r"([A-Z]+)", str(cell_reference or "").upper())
    if not letters:
        return 10**9

    number = 0
    for char in letters.group(1):
        number = number * 26 + (ord(char) - 64)
    return number


def _clean_number(value):
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def set_numeric_cell(row_element, row_number, column_letter, value):
    target_ref = f"{column_letter}{row_number}"
    target_cell = None
    cells = row_element.xpath('./*[local-name()="c"]')

    for cell in cells:
        if cell.get("r") == target_ref:
            target_cell = cell
            break

    if target_cell is None:
        namespace = etree.QName(row_element).namespace
        target_cell = etree.Element(f"{{{namespace}}}c", r=target_ref)
        target_col = column_number(target_ref)
        inserted = False

        for existing in cells:
            if column_number(existing.get("r", "")) > target_col:
                existing.addprevious(target_cell)
                inserted = True
                break
        if not inserted:
            row_element.append(target_cell)

    for child in list(target_cell):
        target_cell.remove(child)
    target_cell.attrib.pop("t", None)

    namespace = etree.QName(target_cell).namespace
    value_node = etree.SubElement(target_cell, f"{{{namespace}}}v")
    value = _clean_number(value)
    value_node.text = str(value) if isinstance(value, int) else format(float(value), ".15g")


__all__ = [
    "load_shared_strings",
    "read_cell_text",
    "find_sheet_xml_path",
    "column_number",
    "set_numeric_cell",
]
