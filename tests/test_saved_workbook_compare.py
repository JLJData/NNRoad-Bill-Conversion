import tempfile
import unittest
from pathlib import Path
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook
from bill_validation.saved_workbook_compare import compare_saved, NS


def write_fixture(path, value, kind='n', missing=False):
    wb = Workbook()
    wb.active.title = 'PN'
    wb.active['E14'] = '=432540000/7'
    wb.active['A1'] = '0001234567890123456'
    source = BytesIO()
    wb.save(source)
    wb.close()
    source.seek(0)
    with ZipFile(source) as archive, ZipFile(path, 'w') as output:
        for name in archive.namelist():
            raw = archive.read(name)
            if name == 'xl/worksheets/sheet1.xml':
                root = ET.fromstring(raw)
                cell = next(c for c in root.iter(f'{NS}c') if c.get('r') == 'E14')
                cell.set('t', kind)
                cache = cell.find(f'{NS}v')
                if missing:
                    cell.remove(cache)
                else:
                    cache.text = value
                raw = ET.tostring(root)
            output.writestr(name, raw)


class SavedWorkbookCompareTests(unittest.TestCase):
    def test_exact_xml_values_and_failures(self):
        with tempfile.TemporaryDirectory() as directory:
            baseline, actual = Path(directory) / 'wps.xlsx', Path(directory) / 'converted.xlsx'
            write_fixture(baseline, '61791428.5714286')
            write_fixture(actual, '61791428.57142857')
            self.assertEqual(compare_saved(baseline, actual)['status'], 'different')
            write_fixture(actual, '61791428.5714286')
            self.assertEqual(compare_saved(baseline, actual)['differences'], [])
            write_fixture(actual, '61791428.571429')
            report = compare_saved(baseline, actual)
            self.assertEqual(report['differences'][0]['cell'], 'E14')
            self.assertEqual(report['differences'][0]['actual']['raw'], '61791428.571429')
            write_fixture(actual, '', missing=True)
            self.assertEqual(compare_saved(baseline, actual)['differences'][0]['reason'], 'invalid-cache')
            for path in (baseline, actual):
                write_fixture(path, '#DIV/0!', 'e')
            self.assertEqual(compare_saved(baseline, actual)['status'], 'different')
            write_fixture(baseline, '', 'str')
            write_fixture(actual, '', 'str')
            self.assertEqual(compare_saved(baseline, actual)['differences'], [])
            write_fixture(actual, '0')
            self.assertEqual(compare_saved(baseline, actual)['status'], 'different')


if __name__ == '__main__':
    unittest.main()
