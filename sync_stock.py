import json
import os
import posixpath
import re
import zipfile
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

import msal
import requests
from lxml import etree
from openpyxl import load_workbook

GRAPH = "https://graph.microsoft.com/v1.0"

HOSTNAME = "vikodacomvn.sharepoint.com"
SITE_PATH = "/sites/Planning"

SOURCE_ACTUAL_PATH = (
    "Tinh san xuat Mua hang 2027/"
    "Ton thuc te/"
    "Bao cao ton thuc te hien tai.xlsx"
)

SOURCE_FACTORY_PATH = (
    "Tinh san xuat Mua hang 2027/"
    "Ton He thong/"
    "Ton Nha May/"
    "NXT_Vikoda.xlsm"
)

DEST_PATH = "Tinh san xuat Mua hang 2027/Sắp kế hoạch.xlsx"
DEST_SHEET = "Ton_kho"

STATE_FILE = Path("state.json")
CODE_PATTERN = re.compile(r"^\d{6,}$")


def get_access_token():
    tenant_id = os.environ["MS_TENANT_ID"]
    client_id = os.environ["MS_CLIENT_ID"]
    client_secret = os.environ["MS_CLIENT_SECRET"]

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )

    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    token = result.get("access_token")
    if not token:
        raise RuntimeError(
            "Không lấy được Microsoft Graph access token: "
            + json.dumps(result, ensure_ascii=False)
        )

    return token


class GraphClient:
    def __init__(self, token):
        self.session = requests.Session()
        self.session.headers.update(
            {"Authorization": f"Bearer {token}"}
        )

    def get_json(self, url, params=None):
        response = self.session.get(
            url,
            params=params,
            timeout=60,
        )
        self._raise(response)
        return response.json()

    def get_site_id(self):
        url = f"{GRAPH}/sites/{HOSTNAME}:{SITE_PATH}"
        return self.get_json(url, {"$select": "id"})["id"]

    def get_default_drive_id(self, site_id):
        url = f"{GRAPH}/sites/{site_id}/drive"
        return self.get_json(url, {"$select": "id"})["id"]

    def get_item_by_path(self, drive_id, file_path):
        encoded = quote(file_path, safe="/")
        url = f"{GRAPH}/drives/{drive_id}/root:/{encoded}"
        return self.get_json(
            url,
            {"$select": "id,name,eTag,size,lastModifiedDateTime"},
        )

    def download_file(self, drive_id, item_id):
        url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/content"
        response = self.session.get(
            url,
            timeout=120,
            allow_redirects=True,
        )
        self._raise(response)
        return response.content

    def upload_file(self, drive_id, item_id, content, expected_etag):
        url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/content"

        headers = {
            "Content-Type": (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            "If-Match": expected_etag,
        }

        response = self.session.put(
            url,
            headers=headers,
            data=content,
            timeout=180,
        )

        if response.status_code == 412:
            raise RuntimeError(
                "File đích vừa thay đổi trong lúc workflow đang chạy. "
                "Dừng để tránh ghi đè dữ liệu mới của người dùng."
            )

        self._raise(response)
        return response.json()

    @staticmethod
    def _raise(response):
        if response.ok:
            return

        try:
            detail = response.json()
        except Exception:
            detail = response.text

        raise RuntimeError(
            f"Microsoft Graph lỗi {response.status_code}: {detail}"
        )


def load_state():
    if not STATE_FILE.exists():
        return {}

    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}

    if "sources" not in data and data.get("source_etag"):
        return {
            "sources": {
                "actual_stock": data["source_etag"],
            }
        }

    return data


def save_state(source_etags):
    STATE_FILE.write_text(
        json.dumps(
            {"sources": source_etags},
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )


def normalize_code(value):
    if value is None:
        return None

    if isinstance(value, float) and value.is_integer():
        value = int(value)

    text = str(value).strip()
    return text if CODE_PATTERN.fullmatch(text) else None


def to_number(value, cell_name):
    if value is None or value == "":
        return 0

    if isinstance(value, bool):
        raise ValueError(f"{cell_name} chứa TRUE/FALSE, không phải số.")

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip().replace(",", "")
    try:
        return float(text)
    except ValueError as exc:
        raise ValueError(
            f"{cell_name} có giá trị không phải số: {value!r}"
        ) from exc


def clean_number(value):
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def read_actual_stock(source_bytes):
    workbook = load_workbook(
        BytesIO(source_bytes),
        data_only=True,
        read_only=True,
    )

    worksheet = workbook.worksheets[0]
    print(f"[Tồn thực tế] Sheet nguồn: {worksheet.title}")

    result = {}

    for row in range(1, worksheet.max_row + 1):
        code = normalize_code(
            worksheet.cell(row=row, column=3).value
        )
        if not code:
            continue

        if code in result:
            raise RuntimeError(
                f"[Tồn thực tế] Mã {code} bị lặp trong cột C."
            )

        value_n = to_number(
            worksheet.cell(row=row, column=14).value,
            f"N{row}",
        )
        value_o = to_number(
            worksheet.cell(row=row, column=15).value,
            f"O{row}",
        )

        result[code] = clean_number(value_n + value_o)

    if not result:
        raise RuntimeError(
            "[Tồn thực tế] Không đọc được mã sản phẩm nào từ cột C."
        )

    print(
        f"[Tồn thực tế] Đọc được {len(result)} mã, "
        "đích Ton_kho cột D = N + O."
    )
    return result


def read_factory_stock(source_bytes):
    workbook = load_workbook(
        BytesIO(source_bytes),
        data_only=True,
        read_only=True,
        keep_vba=True,
    )

    if "Sheet1" not in workbook.sheetnames:
        raise RuntimeError(
            "[Tồn nhà máy] Không tìm thấy Sheet1 trong NXT_Vikoda.xlsm."
        )

    worksheet = workbook["Sheet1"]
    print("[Tồn nhà máy] Sheet nguồn: Sheet1")

    result = {}

    for row in range(1, worksheet.max_row + 1):
        code = normalize_code(
            worksheet.cell(row=row, column=2).value
        )
        if not code:
            continue

        if code in result:
            raise RuntimeError(
                f"[Tồn nhà máy] Mã {code} bị lặp trong cột B."
            )

        value_l = to_number(
            worksheet.cell(row=row, column=12).value,
            f"L{row}",
        )

        result[code] = clean_number(value_l)

    if not result:
        raise RuntimeError(
            "[Tồn nhà máy] Không đọc được mã sản phẩm nào từ cột B."
        )

    print(
        f"[Tồn nhà máy] Đọc được {len(result)} mã, "
        "đích Ton_kho cột E = L."
    )
    return result


def load_shared_strings(archive):
    path = "xl/sharedStrings.xml"
    if path not in archive.namelist():
        return []

    root = etree.fromstring(archive.read(path))
    result = []

    for si in root.xpath('//*[local-name()="si"]'):
        result.append("".join(si.itertext()))

    return result


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
    workbook_root = etree.fromstring(
        archive.read("xl/workbook.xml")
    )

    relationship_ns = (
        "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships"
    )

    relationship_id = None

    for sheet in workbook_root.xpath(
        '//*[local-name()="sheet"]'
    ):
        if sheet.get("name") == sheet_name:
            relationship_id = sheet.get(
                f"{{{relationship_ns}}}id"
            )
            break

    if not relationship_id:
        raise RuntimeError(
            f"Không tìm thấy sheet {sheet_name!r} trong file đích."
        )

    rels_root = etree.fromstring(
        archive.read("xl/_rels/workbook.xml.rels")
    )

    target = None

    for rel in rels_root.xpath(
        '//*[local-name()="Relationship"]'
    ):
        if rel.get("Id") == relationship_id:
            target = rel.get("Target")
            break

    if not target:
        raise RuntimeError(
            f"Không xác định được XML của sheet {sheet_name!r}."
        )

    if target.startswith("/"):
        return target.lstrip("/")

    return posixpath.normpath(
        posixpath.join("xl", target)
    )


def column_number(cell_reference):
    letters = re.match(r"([A-Z]+)", cell_reference)
    if not letters:
        return 10**9

    number = 0
    for char in letters.group(1):
        number = number * 26 + (ord(char) - 64)

    return number


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
        target_cell = etree.Element(
            f"{{{namespace}}}c",
            r=target_ref,
        )

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
    value_node = etree.SubElement(
        target_cell,
        f"{{{namespace}}}v",
    )

    if isinstance(value, int):
        value_node.text = str(value)
    else:
        value_node.text = format(float(value), ".15g")


def patch_destination_workbook(dest_bytes, column_sources):
    source_buffer = BytesIO(dest_bytes)
    output_buffer = BytesIO()

    with zipfile.ZipFile(source_buffer, "r") as source_zip:
        sheet_path = find_sheet_xml_path(
            source_zip,
            DEST_SHEET,
        )

        shared_strings = load_shared_strings(source_zip)

        sheet_root = etree.fromstring(
            source_zip.read(sheet_path)
        )

        matched = {
            column: 0 for column in column_sources
        }
        missing = {
            column: [] for column in column_sources
        }
        seen_dest_codes = set()

        rows = sheet_root.xpath(
            '//*[local-name()="sheetData"]/*[local-name()="row"]'
        )

        for row_element in rows:
            row_number_text = row_element.get("r")
            if not row_number_text:
                continue

            row_number = int(row_number_text)

            code = None
            for cell in row_element.xpath(
                './*[local-name()="c"]'
            ):
                if cell.get("r") == f"A{row_number}":
                    code = normalize_code(
                        read_cell_text(
                            cell,
                            shared_strings,
                        )
                    )
                    break

            if not code:
                continue

            if code in seen_dest_codes:
                raise RuntimeError(
                    f"Mã {code} bị lặp trong cột A sheet {DEST_SHEET}."
                )
            seen_dest_codes.add(code)

            for column, source in column_sources.items():
                values = source["values"]
                if code not in values:
                    missing[column].append(code)
                    continue

                value = values[code]
                set_numeric_cell(
                    row_element,
                    row_number,
                    column,
                    value,
                )
                matched[column] += 1
                print(
                    f"{code}: {DEST_SHEET}!{column}{row_number} = {value} "
                    f"({source['label']})"
                )

        for column, source in column_sources.items():
            if matched[column] == 0:
                raise RuntimeError(
                    f"[{source['label']}] Không có mã nào khớp "
                    f"với cột A sheet {DEST_SHEET}."
                )

            print(
                f"[{source['label']}] Đã cập nhật {matched[column]} mã "
                f"vào cột {column}."
            )

            if missing[column]:
                print(
                    f"[{source['label']}] Không tìm thấy trong nguồn, "
                    "giữ nguyên giá trị cũ: "
                    + ", ".join(missing[column])
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

    return output_buffer.getvalue()


def main():
    token = get_access_token()
    graph = GraphClient(token)

    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    print("Đã kết nối site SharePoint Planning.")

    source_actual_item = graph.get_item_by_path(
        drive_id,
        SOURCE_ACTUAL_PATH,
    )
    source_factory_item = graph.get_item_by_path(
        drive_id,
        SOURCE_FACTORY_PATH,
    )

    current_etags = {
        "actual_stock": source_actual_item["eTag"],
        "factory_stock": source_factory_item["eTag"],
    }

    old_etags = load_state().get("sources", {})

    changed_sources = [
        key
        for key, etag in current_etags.items()
        if old_etags.get(key) != etag
    ]

    print(
        "[Tồn thực tế] Sửa lần cuối:",
        source_actual_item.get("lastModifiedDateTime"),
    )
    print(
        "[Tồn nhà máy] Sửa lần cuối:",
        source_factory_item.get("lastModifiedDateTime"),
    )

    if not changed_sources:
        print("Không có file nguồn nào thay đổi. Kết thúc.")
        return

    print(
        "Nguồn thay đổi:",
        ", ".join(changed_sources),
    )

    dest_item = graph.get_item_by_path(
        drive_id,
        DEST_PATH,
    )

    actual_bytes = graph.download_file(
        drive_id,
        source_actual_item["id"],
    )
    factory_bytes = graph.download_file(
        drive_id,
        source_factory_item["id"],
    )
    dest_bytes = graph.download_file(
        drive_id,
        dest_item["id"],
    )

    actual_stock = read_actual_stock(actual_bytes)
    factory_stock = read_factory_stock(factory_bytes)

    updated_dest_bytes = patch_destination_workbook(
        dest_bytes,
        {
            "D": {
                "label": "Tồn thực tế",
                "values": actual_stock,
            },
            "E": {
                "label": "Tồn nhà máy",
                "values": factory_stock,
            },
        },
    )

    result = graph.upload_file(
        drive_id,
        dest_item["id"],
        updated_dest_bytes,
        expected_etag=dest_item["eTag"],
    )

    print(
        "Upload thành công:",
        result.get("name", "Sắp kế hoạch.xlsx"),
    )

    save_state(current_etags)
    print("SYNC THÀNH CÔNG.")


if __name__ == "__main__":
    main()
