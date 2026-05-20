"""Stamp a project's Project Setup tab in a fresh copy of the master register.

One-off setup helper — run once per project to mint the master register, then
upload it to /Projects/<Suburb> - <Address>/Variations/ in Dropbox.

Example:
    python3 scripts/setup_project_register.py \\
        --project-short "137a" \\
        --project-name "137 Annandale Street, Annandale" \\
        --client "Jon and Stacey Day" \\
        --contract-date 2025-08-20 \\
        --contract-sum 2997386 \\
        --dropbox-folder "Annandale - 137 Annandale Street" \\
        --out working/137-annandale-variation-register.xlsx
"""
from __future__ import annotations

import argparse
import datetime as dt
import shutil
from pathlib import Path

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = REPO_ROOT / "templates" / "master_register_template.xlsx"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--project-short", required=True,
                    help='Short project code, e.g. "137a"')
    ap.add_argument("--project-name", required=True,
                    help='Full project name for display, e.g. "137 Annandale Street, Annandale"')
    ap.add_argument("--client", required=True,
                    help='Client name(s), e.g. "Jon and Stacey Day"')
    ap.add_argument("--contract-date", required=True,
                    help="ISO date YYYY-MM-DD")
    ap.add_argument("--contract-sum", required=True, type=float,
                    help="Original contract sum, inc GST, as a number")
    ap.add_argument("--dropbox-folder", required=True,
                    help='Dropbox project folder name under /Projects, '
                         'e.g. "Annandale - 137 Annandale Street"')
    ap.add_argument("--builder", default="Gold and Eagle Constructions trading as Renovate 8")
    ap.add_argument("--abn", default="80 622 077 472")
    ap.add_argument("--licence", default="139701C")
    ap.add_argument("--out", required=True, help="Output xlsx path")
    args = ap.parse_args()

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(TEMPLATE, out)

    wb = load_workbook(out)

    # Project Setup tab — single-column key/value layout (col A keys, col B values)
    ws = wb["Project Setup"]
    setup = {
        "Project short code": args.project_short,
        "Project name": args.project_name,
        "Client": args.client,
        "Builder": args.builder,
        "ABN": args.abn,
        "Builder licence": args.licence,
        "Contract date": dt.date.fromisoformat(args.contract_date),
        "Original contract sum (inc GST)": args.contract_sum,
        "Dropbox project folder": f"/Projects/{args.dropbox_folder}",
        "Variations folder": f"/Projects/{args.dropbox_folder}/Variations",
        "Set up on": dt.date.today(),
    }
    # Write into the Project Setup tab. We assume the template's column A holds
    # the labels and column B is empty for values. We scan for matching labels.
    existing_labels = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str):
            existing_labels[v.strip().lower()] = r

    next_blank = ws.max_row + 1
    for key, val in setup.items():
        row = existing_labels.get(key.lower())
        if row is None:
            row = next_blank
            ws.cell(row=row, column=1, value=key)
            next_blank += 1
        ws.cell(row=row, column=2, value=val)

    wb.save(out)
    print(f"Created: {out}")
    print(f"  Project: {args.project_name}")
    print(f"  Client:  {args.client}")
    print(f"  Sum:     ${args.contract_sum:,.2f} inc GST")
    print(f"  Dropbox: /Projects/{args.dropbox_folder}/Variations/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
