from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from app import _group_rows, _parse_workbook
from xlsx_tools import read_xlsx_with_merges, write_xlsx


class ExcelHeaderMergeTests(unittest.TestCase):
    def test_reads_and_preserves_header_merge_ranges(self) -> None:
        merge_ranges = [(1, 0, 2, 0), (1, 1, 1, 2)]

        with tempfile.TemporaryDirectory() as temp_name:
            workbook_path = Path(temp_name) / "merged-headers.xlsx"
            write_xlsx(
                workbook_path,
                [["商品名称", "商品属性", ""], ["", "品牌", "型号"]],
                [["商品A", "无品牌", "M1"]],
                merge_header_groups=True,
                merge_ranges=merge_ranges,
            )

            rows, parsed_ranges = read_xlsx_with_merges(workbook_path)

            self.assertEqual(rows[0], ["商品名称", "商品属性"])
            self.assertEqual(rows[1], ["", "品牌", "型号"])
            self.assertEqual(parsed_ranges, merge_ranges)

            with zipfile.ZipFile(workbook_path) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")

        self.assertIn('<mergeCells count="2">', sheet_xml)
        self.assertEqual(sheet_xml.count("<mergeCell "), 2)
        self.assertIn('<mergeCell ref="A1:A2"/>', sheet_xml)
        self.assertIn('<mergeCell ref="B1:C1"/>', sheet_xml)

    def test_uses_merged_header_values_for_product_column_detection(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            workbook_path = Path(temp_name) / "vertical-header.xlsx"
            write_xlsx(
                workbook_path,
                [["商品名称", "商品属性", ""], ["", "品牌", "型号"]],
                [["商品A", "无品牌", "M1"], ["", "无品牌", "M2"]],
                merge_ranges=[(1, 0, 2, 0), (1, 1, 1, 2)],
            )

            header_rows, headers, rows, header_merges = _parse_workbook(workbook_path)

        self.assertEqual(header_rows[1][0], "")
        self.assertEqual(headers[0], "商品名称")
        self.assertEqual(header_merges, [(1, 0, 2, 0), (1, 1, 1, 2)])
        self.assertEqual(list(_group_rows(headers, rows)), ["商品A"])


if __name__ == "__main__":
    unittest.main()
