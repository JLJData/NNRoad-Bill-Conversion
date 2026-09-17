"""Offline code collection reader: python -m bill_validation --help."""
import argparse
import json
import sys
from pathlib import Path
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from . import read_code_collection


def main(argv=None):
    parser = argparse.ArgumentParser(description="Read code collection; never modify the workbook or authorize PASS.")
    for name in ("workbook", "schema", "binding", "manifest"):
        parser.add_argument("--" + name, required=True, type=Path)
    parser.add_argument("--snapshot", type=Path, help="Trusted, successful recalculation snapshot bound to result SHA-256")
    parser.add_argument("--allow-draft", action="store_true", help="Allow offline use of draft configuration")
    parser.add_argument("--output", type=Path, help="JSON output; default stdout. Existing files are never overwritten.")
    args = parser.parse_args(argv)
    try:
        inputs = [args.workbook, args.schema, args.binding, args.manifest, args.snapshot]
        if args.output and args.output.resolve() in {p.resolve() for p in inputs if p}:
            raise ValueError("Output must not overwrite an input")
        def read_json(path):
            return json.loads(path.read_text(encoding="utf-8-sig")) if path else None
        result = read_code_collection(args.workbook, read_json(args.schema), read_json(args.binding),
                                      read_json(args.manifest), calculation_snapshot=read_json(args.snapshot),
                                      allow_draft=args.allow_draft)
        payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
        if args.output:
            with args.output.open("x", encoding="utf-8") as out:
                out.write(payload)
        else:
            print(payload, end="")
        return 0 if result["readiness"] == "ready_for_comparison" else 2
    except (ValueError, OSError, KeyError, TypeError, BadZipFile, InvalidFileException) as exc:
        print("Collection reader: " + str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
