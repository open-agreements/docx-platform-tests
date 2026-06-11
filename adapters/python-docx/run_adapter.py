#!/usr/bin/env python3
"""docx-platform-tests adapter protocol v1 entrypoint for python-docx.

Capability map (see README.md):
- replaceFirstTextOccurrence: intra-run replacement via the library's
  paragraph/run API. A match that spans runs is declined (exit 2) rather
  than approximated with adapter-side algorithms.
- acceptAllTrackedChanges / rejectAllTrackedChanges and everything else:
  exit 2 unsupported -- python-docx has no revision API.
"""
import argparse
import json
import sys

PROTOCOL_VERSION = "1"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--protocol-version", required=True)
    parser.add_argument("--operation", required=True)
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    if args.protocol_version != PROTOCOL_VERSION:
        print(
            f"python-docx adapter speaks protocol v{PROTOCOL_VERSION}, "
            f"got {args.protocol_version}"
        )
        return 3

    with open(args.operation, encoding="utf-8") as f:
        operation = json.load(f)
    name = operation.get("operationName")

    if name != "replaceFirstTextOccurrence":
        if name in ("acceptAllTrackedChanges", "rejectAllTrackedChanges"):
            print("python-docx has no tracked-changes (revision) API")
        else:
            print(f"python-docx adapter does not implement operation '{name}'")
        return 2

    import docx  # imported late so unsupported paths need no library

    find_text = operation["findText"]
    replace_text = operation["replaceText"]
    document = docx.Document(args.input)

    for paragraph in document.paragraphs:
        if find_text not in paragraph.text:
            continue
        for run in paragraph.runs:
            if find_text in run.text:
                run.text = run.text.replace(find_text, replace_text, 1)
                document.save(args.output)
                return 0
        print(
            "match spans run boundaries; the python-docx adapter only "
            "performs intra-run replacement (glue, not algorithms)"
        )
        return 2

    print(f"findText not present in any paragraph: {find_text!r}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
