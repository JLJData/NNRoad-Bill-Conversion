# -*- coding: utf-8 -*-
import unittest
from unittest.mock import MagicMock

from bill_convert.column_cell_map import (
    apply_column_cell_map_to_record,
    parse_column_cell_map,
)


class ColumnCellMapTest(unittest.TestCase):
    def test_parse_letters_and_ck7(self):
        m = {"columnCellMap": [{"sourceColumn": "ck7", "target": "Medical Insurance"}]}
        rows = parse_column_cell_map(m)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["sourceColumn"], "CK")
        self.assertEqual(rows[0]["sourceCol"], 89)
        self.assertEqual(rows[0]["target"], "Medical Insurance")

    def test_apply_does_not_overwrite_by_default(self):
        ws = MagicMock()
        cell = MagicMock()
        cell.value = 42
        ws.cell.return_value = cell
        record = {"Target Col": 1}
        apply_column_cell_map_to_record(
            record,
            8,
            ws,
            {"columnCellMap": [{"sourceColumn": "CK", "target": "Target Col"}]},
        )
        self.assertEqual(record["Target Col"], 1)
        ws.cell.assert_not_called()

    def test_apply_writes_when_empty(self):
        ws = MagicMock()
        cell = MagicMock()
        cell.value = 99.5
        ws.cell.return_value = cell
        record: dict = {}
        apply_column_cell_map_to_record(
            record,
            8,
            ws,
            {"columnCellMap": [{"sourceColumn": "CK", "target": "Target Col"}]},
            coerce_number=lambda v: float(v) if v is not None else None,
        )
        self.assertEqual(record["Target Col"], 99.5)
        ws.cell.assert_called_with(8, 89)

    def test_empty_mapping_no_op(self):
        self.assertEqual(parse_column_cell_map({}), [])
        self.assertEqual(parse_column_cell_map({"columnCellMap": []}), [])


if __name__ == "__main__":
    unittest.main()
