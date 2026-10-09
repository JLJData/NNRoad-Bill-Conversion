"""Compare saved PN values with an Excel/WPS-saved baseline without recalculating.

Usage: python -m bill_validation.saved_workbook_compare baseline.xlsx converted.xlsx --output report.json
Only verifies saved cell values, not baseline freshness or PDF layout.
"""
from __future__ import annotations

import argparse
from decimal import Decimal
import hashlib
from io import BytesIO
import json
from pathlib import Path
import posixpath
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def read_saved(path, sheet_name='PN'):
    raw = Path(path).read_bytes()
    with ZipFile(BytesIO(raw)) as archive:
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        targets = {item.get('Id'): item.get('Target')
                   for item in ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))}
        sheet = next((s for s in workbook.findall(f'{NS}sheets/{NS}sheet')
                      if s.get('name', '').casefold() == sheet_name.casefold()), None)
        if sheet is None:
            raise ValueError(f'{Path(path).name}: missing sheet {sheet_name}')
        target = targets[sheet.get(f'{REL}id')]
        target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            strings = [''.join(t.text or '' for t in item.iter(f'{NS}t'))
                       for item in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
        cells = {}
        for cell in ET.fromstring(archive.read(target)).iter(f'{NS}c'):
            formula, value = cell.find(f'{NS}f'), cell.find(f'{NS}v')
            kind, text = cell.get('t', 'n'), '' if value is None or value.text is None else value.text
            if kind == 's' and value is not None:
                kind, text = 'str', strings[int(text)]
            elif kind == 'inlineStr':
                kind, text = 'str', ''.join(t.text or '' for t in cell.iter(f'{NS}t'))
            if formula is not None and (value is None or (not text and kind != 'str')):
                kind = 'missing-cache'
            elif value is None and kind != 'str':
                continue
            cells[cell.get('r')] = {'type': kind, 'raw': text,
                                  'formula': None if formula is None else formula.text or ''}
        properties = workbook.find(f'{NS}workbookPr')
        date1904 = properties is not None and properties.get('date1904') in ('1', 'true')
    return {'sha256': hashlib.sha256(raw).hexdigest(), 'date1904': date1904, 'cells': cells}


def canonical(cell):
    if cell is None:
        return None
    kind, raw = cell['type'], cell['raw']
    if kind == 'n':
        value = Decimal(raw)
        if not value.is_finite():
            raise ValueError('Non-finite numeric cache')
        return kind, value
    return kind, raw


def compare_saved(reference, actual, sheet='PN'):
    left, right = read_saved(reference, sheet), read_saved(actual, sheet)
    differences = []
    if left['date1904'] != right['date1904']:
        differences.append({'cell': None, 'reason': 'date-system-mismatch'})
    for ref in sorted(left['cells'].keys() | right['cells'].keys()):
        a, b = left['cells'].get(ref), right['cells'].get(ref)
        invalid = any(c and c['type'] in ('e', 'missing-cache') for c in (a, b))
        if invalid or canonical(a) != canonical(b):
            differences.append({'cell': ref, 'reason': 'invalid-cache' if invalid else 'value-or-type-mismatch',
                                'reference': a, 'actual': b})
    return {'status': 'different' if differences else 'matched-saved-values', 'sheet': sheet,
            'policy': 'stored-decimal-exact; exact-text-and-types',
            'referenceSha256': left['sha256'], 'actualSha256': right['sha256'],
            'comparedCellCount': len(left['cells'].keys() | right['cells'].keys()),
            'differences': differences,
            'scope': 'Saved values only; baseline freshness and PDF layout require separate verification.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('reference')
    parser.add_argument('actual')
    parser.add_argument('--sheet', default='PN')
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    report = compare_saved(args.reference, args.actual, args.sheet)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        args.output.write_text(text + '\n', encoding='utf-8')
    else:
        print(text)
    return 1 if report['differences'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
