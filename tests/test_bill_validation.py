import copy
import hashlib
import json
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from io import BytesIO
from zipfile import ZipFile

from openpyxl import Workbook

from bill_validation import ValidationError, read_code_collection, validate_collection
from bill_validation.__main__ import main


class CollectionReaderTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "result.xlsx"
        self.schema = {
            "formatVersion": 0, "schemaId": "pay", "schemaVersion": "1",
            "status": "draft_not_enabled", "automaticPassEnabled": False,
            "decimalEncoding": "string", "blankIsZero": False,
            "fields": [{"field": f, "type": "decimal", "unit": "currency"} for f in ("salary", "allowance")],
        }
        self.binding = {
            "formatVersion": 0, "schemaId": "pay", "schemaVersion": "1",
            "status": "draft_not_enabled", "automaticPassEnabled": False,
            "template": {"sha256": "a" * 64}, "resultSheet": "Collection", "headerRow": 1,
            "entity": {"latinNameColumn": "A"},
            "fields": [{"field": "salary", "columnHint": "B", "expectedHeader": "Basic Salary"},
                       {"field": "allowance", "columnHint": "C", "expectedHeader": "Allowance"}],
        }
        self.manifest = {
            "templateSha256": "a" * 64, "period": "2026-03", "currency": "TWD",
            "employees": [{"row": 2, "entityKey": "e1", "expectedName": "Alice"},
                          {"row": 3, "entityKey": "e2", "expectedName": "Bob"}],
        }
        self.write([['Name', 'Basic Salary', 'Allowance'], ['Alice', 100, 0], ['Bob', -2.50, None]])

    def write(self, rows):
        wb = Workbook()
        ws = wb.active
        ws.title = "Collection"
        for row in rows:
            ws.append(row)
        wb.save(self.path)
        wb.close()
        self.manifest['resultSha256'] = hashlib.sha256(self.path.read_bytes()).hexdigest()

    def read(self, **kwargs):
        return read_code_collection(self.path, self.schema, self.binding, self.manifest, allow_draft=True, **kwargs)

    def snapshot(self, cells):
        return {"status": "succeeded", "resultSha256": self.manifest['resultSha256'],
                "engine": "test-fixture", "cells": cells}

    def test_zero_missing_signed_decimal_and_provenance(self):
        before = self.path.read_bytes()
        result = self.read()
        a, b = result['records']
        self.assertEqual(a['fields'][1]['value'], '0')
        self.assertEqual(b['fields'][0]['value'], '-2.5')
        self.assertEqual(b['fields'][1]['status'], 'missing')
        self.assertIsNone(b['fields'][1]['value'])
        self.assertEqual(a['fields'][0]['source']['cell'], 'B2')
        self.assertEqual(result['readiness'], 'review_required')
        self.assertEqual(self.path.read_bytes(), before)

    def test_unique_headers_can_move_and_rows_can_reorder(self):
        self.write([['Name', 'Allowance', 'Basic\n Salary'], ['Bob', 5, 90], ['Alice', 0, 100]])
        self.manifest['employees'][0]['row'] = 3
        self.manifest['employees'][1]['row'] = 2
        result = self.read()
        self.assertEqual(result['readiness'], 'ready_for_comparison')
        self.assertEqual(result['records'][0]['fields'][0]['value'], '100')
        self.assertEqual(result['records'][0]['fields'][0]['source']['cell'], 'C3')
        self.assertFalse(result['automaticPassEnabled'])

    def test_missing_or_duplicate_headers_are_rejected(self):
        for headers in (['Name', 'Salary', 'Allowance'], ['Name', 'Basic Salary', 'Basic Salary', 'Allowance']):
            with self.subTest(headers=headers):
                self.write([headers, ['Alice', 100, 0], ['Bob', 90, 0]])
                with self.assertRaisesRegex(ValidationError, 'ambiguous header'):
                    self.read()

    def test_unavailable_formula_does_not_become_zero(self):
        self.write([['Name', 'Basic Salary', 'Allowance'], ['Alice', '=10*10', 0], ['Bob', 90, 0]])
        result = self.read()
        self.assertEqual(result['records'][0]['fields'][0]['status'], 'unreadable')
        self.assertIsNone(result['records'][0]['fields'][0]['value'])
        self.assertEqual(result['issues'][0]['code'], 'FORMULA_RESULT_UNAVAILABLE')

    def test_existing_formula_cache_is_not_accepted_as_recalculation(self):
        self.write([['Name', 'Basic Salary', 'Allowance'], ['Alice', '=10*10', 0], ['Bob', 90, 0]])
        buffer = BytesIO()
        with ZipFile(self.path) as original, ZipFile(buffer, 'w') as patched:
            for item in original.infolist():
                data = original.read(item.filename)
                if item.filename == 'xl/worksheets/sheet1.xml':
                    tree = ET.fromstring(data)
                    ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    cell = tree.find(".//m:c[@r='B2']", ns)
                    cell.find('m:v', ns).text = '100'
                    data = ET.tostring(tree, encoding='utf-8')
                patched.writestr(item, data)
        self.path.write_bytes(buffer.getvalue())
        self.manifest['resultSha256'] = hashlib.sha256(self.path.read_bytes()).hexdigest()
        result = self.read()
        self.assertEqual(result['records'][0]['fields'][0]['status'], 'unreadable')
        self.assertEqual(result['issues'][0]['code'], 'FORMULA_RESULT_UNAVAILABLE')

    def test_different_fields_cannot_silently_share_one_column(self):
        self.binding['fields'][1]['expectedHeader'] = 'Basic Salary'
        with self.assertRaisesRegex(ValidationError, 'same column'):
            self.read()

    def test_only_current_successful_snapshot_can_supply_formula_values(self):
        self.write([['Name', 'Basic Salary', 'Allowance'], ['=Other!A1', '=10*10', 0], ['Bob', 90, 0]])
        snap = self.snapshot([
            {'sheet': 'Collection', 'cell': 'A2', 'type': 'string', 'value': 'Alice'},
            {'sheet': 'Collection', 'cell': 'B2', 'type': 'number', 'value': '100.00'},
        ])
        result = self.read(calculation_snapshot=snap)
        self.assertEqual(result['readiness'], 'ready_for_comparison')
        self.assertEqual(result['records'][0]['fields'][0]['value'], '100.00')
        for key, value in [('status', 'failed'), ('resultSha256', 'b' * 64)]:
            invalid = copy.deepcopy(snap)
            invalid[key] = value
            with self.subTest(key=key), self.assertRaises(ValidationError):
                self.read(calculation_snapshot=invalid)

    def test_partial_snapshot_and_formula_errors_require_review(self):
        self.write([['Name', 'Basic Salary', 'Allowance'], ['Alice', '=1/0', '=2+2'], ['Bob', 90, 0]])
        snap = self.snapshot([{'sheet': 'Collection', 'cell': 'B2', 'type': 'error', 'value': '#DIV/0!'}])
        result = self.read(calculation_snapshot=snap)
        self.assertEqual({i['code'] for i in result['issues']}, {'FORMULA_ERROR', 'FORMULA_RESULT_UNAVAILABLE'})

    def test_stale_manifest_or_template_version_is_rejected(self):
        for key in ('resultSha256', 'templateSha256'):
            original = self.manifest[key]
            self.manifest[key] = 'b' * 64
            with self.subTest(key=key), self.assertRaises(ValidationError):
                self.read()
            self.manifest[key] = original

    def test_wrong_employee_or_duplicate_name_requires_review(self):
        self.write([['Name', 'Basic Salary', 'Allowance'], ['Alice', 100, 0], ['Alice', 90, 0]])
        result = self.read()
        self.assertTrue({'IDENTITY_MISMATCH', 'AMBIGUOUS_IDENTITY'} <= {i['code'] for i in result['issues']})

    def test_employee_manifest_cannot_be_guessed(self):
        for entries in ([], [self.manifest['employees'][0]] * 2,
                        [{'row': 999, 'entityKey': 'e1', 'expectedName': 'Alice'}]):
            original = self.manifest['employees']
            self.manifest['employees'] = entries
            with self.subTest(entries=entries), self.assertRaises(ValidationError):
                self.read()
            self.manifest['employees'] = original

    def test_boolean_error_and_formatted_text_are_not_money(self):
        for bad in (True, '#REF!', '1,234.00', 'NaN', 'Infinity'):
            with self.subTest(bad=bad):
                self.write([['Name', 'Basic Salary', 'Allowance'], ['Alice', bad, 0], ['Bob', 90, 0]])
                self.assertEqual(self.read()['records'][0]['fields'][0]['status'], 'unreadable')

    def test_draft_requires_explicit_offline_opt_in(self):
        with self.assertRaisesRegex(ValidationError, 'allow_draft'):
            read_code_collection(self.path, self.schema, self.binding, self.manifest)

    def test_configuration_field_and_version_mismatch(self):
        original = copy.deepcopy(self.binding)
        for mutate in (lambda b: b.update(schemaVersion='old'),
                       lambda b: b['fields'].pop(),
                       lambda b: b['fields'].append(b['fields'][0])):
            self.binding = copy.deepcopy(original)
            mutate(self.binding)
            with self.assertRaises(ValidationError):
                self.read()

    def test_output_contract_rejects_missing_field_and_false_readiness(self):
        result = self.read()
        result['records'][0]['fields'].pop()
        with self.assertRaises(ValidationError):
            validate_collection(result, self.schema)
        result = self.read()
        result['readiness'] = 'ready_for_comparison'
        with self.assertRaises(ValidationError):
            validate_collection(result, self.schema)

    def test_duplicate_snapshot_cell_is_rejected(self):
        item = {'sheet': 'Collection', 'cell': 'B2', 'type': 'number', 'value': '10'}
        with self.assertRaisesRegex(ValidationError, 'Duplicate snapshot'):
            self.read(calculation_snapshot=self.snapshot([item, item]))

    def test_cli_writes_review_json_and_preserves_existing_output(self):
        paths = {}
        for name, value in [('schema', self.schema), ('binding', self.binding), ('manifest', self.manifest)]:
            paths[name] = Path(self.temp.name) / (name + '.json')
            paths[name].write_text(json.dumps(value), encoding='utf-8')
        out = Path(self.temp.name) / 'collection.json'
        argv = ['--workbook', str(self.path), '--allow-draft', '--output', str(out)]
        for name, path in paths.items():
            argv.extend(['--' + name, str(path)])
        self.assertEqual(main(argv), 2)
        self.assertEqual(json.loads(out.read_text(encoding='utf-8'))['readiness'], 'review_required')
        before = out.read_bytes()
        self.assertEqual(main(argv), 1)
        self.assertEqual(out.read_bytes(), before)
        argv[argv.index('--output') + 1] = str(self.path)
        before = self.path.read_bytes()
        self.assertEqual(main(argv), 1)
        self.assertEqual(self.path.read_bytes(), before)


if __name__ == '__main__':
    unittest.main()
