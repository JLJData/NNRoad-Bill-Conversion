"""Source values must survive CODE writing; display precision is a separate policy."""
import importlib
from decimal import Decimal, localcontext, ROUND_HALF_UP
from io import BytesIO
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from ai_collection.template_fill import _dynamic_value
from bill_convert.fixed_value_writes import apply_fixed_value_writes


def saved_xml_values(wb):
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    with ZipFile(output) as archive:
        root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        return {cell.get('r'): cell.findtext(f'{ns}v') for cell in root.iter(f'{ns}c')}


class ConverterNumericPrecisionTests(unittest.TestCase):
    def test_code_and_ai_have_identical_persisted_values_without_rounding_tolerance(self):
        # The independent AI writer uses this numeric decoder. Equal extracted
        # facts must produce identical XML values through both output paths.
        values = ['61791428.57142857', '-61791428.57142857', '0.123456789123456',
                  '0.0000000001', '-0.0000000001', '1.2345678901234567', '0']
        for region in ('cyprus', 'india', 'indonesia', 'italy', 'kyrgyzstan'):
            writer = importlib.import_module(f'profiles.{region}_payroll_calc.convert')._set_cell_value
            with self.subTest(region=region):
                wb = Workbook()
                for row, text in enumerate(values, 1):
                    writer(wb.active.cell(row, 1), float(text))
                    wb.active.cell(row, 2).value = _dynamic_value({'valueType': 'decimal', 'value': text})
                wb.active['C1'] = _dynamic_value({'valueType': 'decimal', 'value': '61791428.5714286'})
                persisted = saved_xml_values(wb)
                wb.close()
                for row in range(1, len(values) + 1):
                    self.assertEqual(persisted[f'A{row}'], persisted[f'B{row}'])
                self.assertEqual(persisted['A1'], '61791428.57142857')
                self.assertNotEqual(persisted['A1'], persisted['C1'])

    def test_fixed_configuration_values_are_not_rounded_by_the_writer(self):
        wb = Workbook()
        wb.active.title = 'PN'
        amount = 61791428.57142857
        records = apply_fixed_value_writes(wb, [{'name': 'Alice'}], {
            'fee': amount, 'fixedValueWrites': [
                {'valueKey': 'fee', 'sheet': 'PN', 'column': 1, 'dataStartRow': 1}]})
        self.assertEqual(wb.active['A1'].value, amount)
        self.assertEqual(records[0]['value'], amount)
        self.assertEqual(saved_xml_values(wb)['A1'], '61791428.57142857')
        wb.close()

    def test_source_precision_survives_xlsx_roundtrip(self):
        for region in ('cyprus', 'india', 'indonesia', 'italy', 'kyrgyzstan'):
            writer = importlib.import_module(f'profiles.{region}_payroll_calc.convert')._set_cell_value
            with self.subTest(region=region):
                wb = Workbook()
                values = [432540000 / 7, -432540000 / 7, 0.123456789123456,
                          1e-10, -1e-10, 0, '0001234567890123456', '=1/7']
                for row, value in enumerate(values, 1):
                    writer(wb.active.cell(row, 1), value)
                    self.assertEqual(wb.active.cell(row, 1).value, value)
                output = BytesIO()
                wb.save(output)
                wb.close()
                output.seek(0)
                saved = load_workbook(output, data_only=False)
                try:
                    for row, value in enumerate(values, 1):
                        actual = saved.active.cell(row, 1).value
                        if isinstance(value, float):
                            # openpyxl serializes 16 significant digits. Check the common
                            # 15-digit policy, while the assertion above checks lossless assignment.
                            with localcontext() as context:
                                context.prec = 15
                                context.rounding = ROUND_HALF_UP
                                self.assertEqual(+Decimal(str(actual)), +Decimal(str(value)))
                        else:
                            self.assertEqual(actual, value)
                finally:
                    saved.close()


if __name__ == '__main__':
    unittest.main()
