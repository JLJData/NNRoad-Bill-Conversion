"""connectSalarySplit must emit cellProvenance coordinates for blue badges."""
import unittest

from bill_convert.connect_salary_split import build_connect_salary_split_cell_writes


class ConnectSalarySplitProvenanceTests(unittest.TestCase):
    def test_matched_employee_emits_basic_housing_transport(self):
        employees = [
            {
                "English Name": "Mohamad Fiazul Huq",
                "Basic Salary": 32000,
                "Housing Allowance": 7700,
                "Transport": 2100,
            }
        ]
        mapping = {
            "connectSalarySplit": {
                "Mohamad Fiazul Huq": {"basic": 32000, "housing": 7700, "transport": 2100},
            }
        }
        header_map = {
            "English Name": 2,
            "Basic Salary": 7,
            "Housing Allowance": 8,
            "Transport": 9,
        }
        cells = build_connect_salary_split_cell_writes(
            employees,
            sheet="UAE-L",
            data_start=8,
            header_map=header_map,
            mapping=mapping,
        )
        self.assertEqual(len(cells), 3)
        self.assertEqual({c["label"] for c in cells}, {"Basic Salary", "Housing Allowance", "Transport"})
        self.assertTrue(all(c["kind"] == "connectSalarySplit" for c in cells))
        self.assertTrue(all(c["row"] == 8 for c in cells))
        self.assertEqual({c["col"] for c in cells}, {7, 8, 9})

    def test_unconfigured_employee_emits_nothing(self):
        employees = [{"English Name": "Someone Else", "Basic Salary": 100}]
        mapping = {"connectSalarySplit": {"Mohamad Fiazul Huq": {"basic": 1}}}
        cells = build_connect_salary_split_cell_writes(
            employees,
            sheet="UAE-L",
            data_start=8,
            header_map={"Basic Salary": 7},
            mapping=mapping,
        )
        self.assertEqual(cells, [])


if __name__ == "__main__":
    unittest.main()
