from __future__ import annotations

import html
import re
import zipfile
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _tag(name: str) -> str:
    return f"{{{XML_NS}}}{name}"


def _col_to_index(cell_ref: str) -> int:
    letters = re.sub(r"[^A-Z]", "", cell_ref.upper())
    value = 0
    for char in letters:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return max(value - 1, 0)


def _column_name(index: int) -> str:
    name = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall(_tag("si")):
        parts = [node.text or "" for node in item.iter(_tag("t"))]
        strings.append("".join(parts))
    return strings


def _first_sheet_path(archive: zipfile.ZipFile) -> str:
    names = set(archive.namelist())
    if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
        return "xl/worksheets/sheet1.xml"

    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    relationships = {
        rel.attrib.get("Id"): rel.attrib.get("Target", "")
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    first_sheet = workbook.find(f"{_tag('sheets')}/{_tag('sheet')}")
    if first_sheet is None:
        return "xl/worksheets/sheet1.xml"

    rel_id = first_sheet.attrib.get(f"{{{REL_NS}}}id")
    target = relationships.get(rel_id, "worksheets/sheet1.xml")
    if target.startswith("/xl/"):
        return target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return "xl/" + target.lstrip("/")


def read_xlsx(path: Path) -> list[list[str]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = _read_shared_strings(archive)
        sheet_path = _first_sheet_path(archive)
        root = ET.fromstring(archive.read(sheet_path))

    rows: list[list[str]] = []
    for row_node in root.findall(f".//{_tag('sheetData')}/{_tag('row')}"):
        row_values: list[str] = []
        for cell in row_node.findall(_tag("c")):
            ref = cell.attrib.get("r", "")
            col_index = _col_to_index(ref)
            while len(row_values) <= col_index:
                row_values.append("")

            value_node = cell.find(_tag("v"))
            raw = value_node.text if value_node is not None else ""
            cell_type = cell.attrib.get("t")

            if cell_type == "s" and raw:
                value = shared_strings[int(raw)]
            elif cell_type == "inlineStr":
                value = "".join(node.text or "" for node in cell.iter(_tag("t")))
            else:
                value = raw or ""

            row_values[col_index] = value.strip()

        while row_values and row_values[-1] == "":
            row_values.pop()
        rows.append(row_values)

    while rows and not any(rows[-1]):
        rows.pop()
    return rows


def write_xlsx(path: Path, header_rows: list[list[str]], rows: Iterable[list[str]]) -> None:
    row_list = list(rows)
    sheet_xml = _build_sheet_xml([*header_rows, *row_list])

    files = {
        "[Content_Types].xml": _content_types_xml(),
        "_rels/.rels": _root_rels_xml(),
        "xl/workbook.xml": _workbook_xml(),
        "xl/_rels/workbook.xml.rels": _workbook_rels_xml(),
        "xl/styles.xml": _styles_xml(),
        "xl/worksheets/sheet1.xml": sheet_xml,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)


def _cell_xml(value: str, row_number: int, col_index: int) -> str:
    ref = f"{_column_name(col_index)}{row_number}"
    escaped = html.escape(str(value or ""))
    return f'<c r="{ref}" t="inlineStr"><is><t>{escaped}</t></is></c>'


def _build_sheet_xml(rows: list[list[str]]) -> str:
    max_columns = max((len(row) for row in rows), default=1)
    last_ref = f"{_column_name(max_columns - 1)}{max(len(rows), 1)}"
    rendered_rows: list[str] = []
    for row_index, row in enumerate(rows, start=1):
        cells = "".join(_cell_xml(value, row_index, col_index) for col_index, value in enumerate(row))
        rendered_rows.append(f'<row r="{row_index}">{cells}</row>')
    sheet_data = "".join(rendered_rows)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{XML_NS}" xmlns:r="{REL_NS}">'
        f'<dimension ref="A1:{last_ref}"/>'
        "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
        "<sheetFormatPr defaultRowHeight=\"18\"/>"
        f"<sheetData>{sheet_data}</sheetData>"
        "</worksheet>"
    )


def _content_types_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""


def _root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def _workbook_xml() -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{XML_NS}" xmlns:r="{REL_NS}">
  <sheets>
    <sheet name="商品数据" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""


def _workbook_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""


def _styles_xml() -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{XML_NS}">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""
