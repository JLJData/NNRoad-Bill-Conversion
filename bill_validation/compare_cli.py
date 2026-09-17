"""Offline deterministic comparison: python -m bill_validation.compare_cli --help."""
import argparse
import json
import sys
from pathlib import Path

from .comparator import compare_collections


def main(argv=None):
    parser = argparse.ArgumentParser(description="Compare code and AI collections without changing workflow state.")
    for name in ("code", "ai", "schema", "documents"):
        parser.add_argument("--" + name, required=True, type=Path)
    parser.add_argument("--output", type=Path, help="JSON output; default stdout. Existing files are never overwritten.")
    args = parser.parse_args(argv)
    try:
        inputs = [args.code, args.ai, args.schema, args.documents]
        if args.output and args.output.resolve() in {path.resolve() for path in inputs}:
            raise ValueError("Output must not overwrite an input")
        def read(path):
            return json.loads(path.read_text(encoding="utf-8-sig"))
        documents = read(args.documents)
        if isinstance(documents, dict):
            documents = documents.get("documents")
        result = compare_collections(read(args.code), read(args.ai), read(args.schema), documents)
        payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
        if args.output:
            with args.output.open("x", encoding="utf-8") as target:
                target.write(payload)
        else:
            print(payload, end="")
        return {"MATCH": 0, "REVIEW_REQUIRED": 2, "MISMATCH": 3}[result["comparisonStatus"]]
    except (ValueError, OSError, KeyError, TypeError) as exc:
        print("Collection comparator: " + str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
