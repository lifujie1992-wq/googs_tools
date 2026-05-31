from __future__ import annotations

import html
import re
import zipfile
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


MergeRange = tuple[int, int, int, int]

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


def _cell_coordinates(cell_ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", cell_ref.upper())
    if match is None:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    col_name, row_number = match.groups()
    return int(row_number), _col_to_index(col_name)


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


def _parse_merge_ref(ref: str) -> MergeRange:
    parts = ref.split(":")
    start_ref = parts[0]
    end_ref = parts[-1]
    start_row, start_col = _cell_coordinates(start_ref)
    end_row, end_col = _cell_coordinates(end_ref)
    return (
        min(start_row, end_row),
        min(start_col, end_col),
        max(start_row, end_row),
        max(start_col, end_col),
    )


def _read_merge_ranges(root: ET.Element) -> list[MergeRange]:
    merge_cells = root.find(_tag("mergeCells"))
    if merge_cells is None:
        return []
    ranges: list[MergeRange] = []
    for merge_cell in merge_cells.findall(_tag("mergeCell")):
        ref = merge_cell.attrib.get("ref", "")
        if not ref:
            continue
        try:
            ranges.append(_parse_merge_ref(ref))
        except ValueError:
            continue
    return ranges


def read_xlsx_with_merges(path: Path) -> tuple[list[list[str]], list[MergeRange]]:
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
    return rows, _read_merge_ranges(root)


def read_xlsx(path: Path) -> list[list[str]]:
    rows, _ = read_xlsx_with_merges(path)
    return rows


def write_xlsx(
    path: Path,
    header_rows: list[list[str]],
    rows: Iterable[list[str]],
    merge_header_groups: bool = False,
    merge_ranges: list[MergeRange] | None = None,
) -> None:
    row_list = list(rows)
    final_merge_ranges = list(merge_ranges or [])
    if merge_header_groups:
        final_merge_ranges = _merge_non_overlapping_ranges(
            final_merge_ranges,
            _header_group_merge_ranges(header_rows),
        )
    sheet_xml = _build_sheet_xml([*header_rows, *row_list], final_merge_ranges)

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


def _header_group_merge_ranges(header_rows: list[list[str]]) -> list[MergeRange]:
    if len(header_rows) < 2:
        return []
    group_row = header_rows[0]
    header_width = max(len(header_rows[0]), len(header_rows[1]))
    ranges: list[MergeRange] = []
    col_index = 0
    while col_index < header_width:
        title = group_row[col_index].strip() if col_index < len(group_row) else ""
        if not title:
            col_index += 1
            continue
        end_index = col_index
        next_index = col_index + 1
        while next_index < header_width:
            next_title = group_row[next_index].strip() if next_index < len(group_row) else ""
            if next_title:
                break
            end_index = next_index
            next_index += 1
        if end_index > col_index:
            ranges.append((1, col_index, 1, end_index))
        col_index = end_index + 1
    return ranges


def _ranges_overlap(first: MergeRange, second: MergeRange) -> bool:
    first_start_row, first_start_col, first_end_row, first_end_col = first
    second_start_row, second_start_col, second_end_row, second_end_col = second
    rows_overlap = first_start_row <= second_end_row and second_start_row <= first_end_row
    columns_overlap = first_start_col <= second_end_col and second_start_col <= first_end_col
    return rows_overlap and columns_overlap


def _merge_non_overlapping_ranges(
    explicit_ranges: list[MergeRange],
    inferred_ranges: list[MergeRange],
) -> list[MergeRange]:
    result = list(dict.fromkeys(explicit_ranges))
    for inferred_range in inferred_ranges:
        if any(_ranges_overlap(inferred_range, existing) for existing in result):
            continue
        result.append(inferred_range)
    return result


def _merge_ref(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    return f"{_column_name(start_col)}{start_row}:{_column_name(end_col)}{end_row}"


def _merge_cells_xml(merge_ranges: list[MergeRange]) -> str:
    if not merge_ranges:
        return ""
    cells = "".join(
        f'<mergeCell ref="{_merge_ref(start_row, start_col, end_row, end_col)}"/>'
        for start_row, start_col, end_row, end_col in merge_ranges
    )
    return f'<mergeCells count="{len(merge_ranges)}">{cells}</mergeCells>'


def _build_sheet_xml(rows: list[list[str]], merge_ranges: list[MergeRange] | None = None) -> str:
    merge_ranges = merge_ranges or []
    max_columns = max(
        max((len(row) for row in rows), default=1),
        max((end_col + 1 for _, _, _, end_col in merge_ranges), default=1),
    )
    max_rows = max(
        max(len(rows), 1),
        max((end_row for _, _, end_row, _ in merge_ranges), default=1),
    )
    last_ref = f"{_column_name(max_columns - 1)}{max_rows}"
    rendered_rows: list[str] = []
    for row_index, row in enumerate(rows, start=1):
        cells = "".join(_cell_xml(value, row_index, col_index) for col_index, value in enumerate(row))
        rendered_rows.append(f'<row r="{row_index}">{cells}</row>')
    sheet_data = "".join(rendered_rows)
    merge_cells = _merge_cells_xml(merge_ranges)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{XML_NS}" xmlns:r="{REL_NS}">'
        f'<dimension ref="A1:{last_ref}"/>'
        "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
        "<sheetFormatPr defaultRowHeight=\"18\"/>"
        f"<sheetData>{sheet_data}</sheetData>"
        f"{merge_cells}"
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
