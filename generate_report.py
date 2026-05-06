#!/usr/bin/env python3
"""Fill a CEZA radiocarbon report template with sample data from a JSON file.

Usage:
    python generate_report.py --template template.docx --data data.json --out report.docx

The JSON file must contain two top-level keys:

  - ``header``: flat dict of header fields (ProjectName, FirstName, LastName,
    organisation, Address1, ZIP, City, email, ProjectInDate, OrderNr).
  - ``samples``: list of dicts, each with: labno, sample_name, c14_age, err,
    d13c, cal_1sigma, cal_2sigma, cn, pct_c, pct_coll, mattype.

See ``data_example.json`` for a complete example.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from docxtpl import DocxTemplate


HEADER_KEYS = {
    "ProjectName", "FirstName", "LastName", "organisation", "Address1",
    "ZIP", "City", "email", "ProjectInDate", "OrderNr",
}
SAMPLE_KEYS = {
    "labno", "sample_name", "c14_age", "err", "d13c",
    "cal_1sigma", "cal_2sigma", "cn", "pct_c", "pct_coll", "mattype",
}


def validate(data: dict) -> None:
    if "header" not in data or "samples" not in data:
        raise SystemExit("JSON must contain 'header' and 'samples' keys.")
    missing = HEADER_KEYS - data["header"].keys()
    if missing:
        raise SystemExit(f"header is missing keys: {sorted(missing)}")
    if not isinstance(data["samples"], list) or not data["samples"]:
        raise SystemExit("'samples' must be a non-empty list.")
    for i, s in enumerate(data["samples"]):
        missing = SAMPLE_KEYS - s.keys()
        if missing:
            raise SystemExit(f"samples[{i}] is missing keys: {sorted(missing)}")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--template", required=True, type=Path)
    ap.add_argument("--data", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    args = ap.parse_args()

    data = json.loads(args.data.read_text(encoding="utf-8"))
    validate(data)

    ctx = {**data["header"], "samples": data["samples"]}
    tpl = DocxTemplate(str(args.template))
    tpl.render(ctx)
    tpl.save(str(args.out))
    print(f"Saved: {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
