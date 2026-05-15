"""Create a new variation: allocate number, create Dropbox folder, stub the
register row.

This is a helper. It does NOT generate the formal Variation Notice
(that's done from the Word template) and it does NOT email anything.

Usage:
    python3 scripts/create_variation.py \
        --project 42SMITH \
        --description "extra-power-points-kitchen" \
        --trigger "Client request"

By default the script writes folders under a LOCAL working directory
(./working/) for testing. To create folders directly in Dropbox via the
Dropbox MCP, run this from the Claude Code harness — Claude will read
the planned actions and execute the Dropbox calls.

For local testing:
    python3 scripts/create_variation.py --project 42SMITH \
        --description "test" --dry-run
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    print("Missing dependency: openpyxl. Run: pip install openpyxl", file=sys.stderr)
    sys.exit(2)


REPO_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = REPO_ROOT / "templates" / "master_register_template.xlsx"
LOCAL_WORKING = REPO_ROOT / "working"


def slugify(s: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", "-", s).strip("-").lower()
    return s[:60] if len(s) > 60 else s


def find_next_var_number(register_path: Path) -> int:
    """Return the next sequential variation number (1-indexed) for the project."""
    if not register_path.exists():
        return 1
    wb = load_workbook(register_path)
    ws = wb["Variations"]
    used = []
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, max_col=1, values_only=True):
        val = row[0]
        if not val:
            continue
        m = re.search(r"VAR-(\d+)\b", str(val))
        if m:
            used.append(int(m.group(1)))
    return (max(used) + 1) if used else 1


def stub_register_row(register_path: Path, *, var_id: str, description: str,
                      trigger: str, dropbox_folder: str) -> int:
    """Append a new RAISED row to the project register. Returns the row number used."""
    wb = load_workbook(register_path)
    ws = wb["Variations"]
    # Find first empty data row (starting at row 4)
    target_row = None
    for r in range(4, ws.max_row + 1):
        if ws.cell(row=r, column=1).value in (None, ""):
            target_row = r
            break
    if target_row is None:
        target_row = ws.max_row + 1

    ws.cell(row=target_row, column=1, value=var_id)                       # A: VAR #
    ws.cell(row=target_row, column=2, value=dt.date.today())              # B: Date Raised
    ws.cell(row=target_row, column=3, value="RAISED")                     # C: Status
    ws.cell(row=target_row, column=4, value="1 — RAISED")                 # D: Phase
    ws.cell(row=target_row, column=5, value=trigger)                      # E: Trigger
    ws.cell(row=target_row, column=6, value=description)                  # F: Short desc
    ws.cell(row=target_row, column=28, value=dropbox_folder)              # AB: Dropbox folder

    wb.save(register_path)
    return target_row


def planned_dropbox_paths(project_code: str, var_id: str, slug: str) -> list[str]:
    base = f"/Renovate 8/Projects/{project_code} - .../03_Variations/{var_id} - {slug}"
    return [
        base,
        f"{base}/01_Source",
        f"{base}/01_Source/photos",
        f"{base}/02_Costing",
        f"{base}/03_Notice",
        f"{base}/04_Approval",
        f"{base}/05_Invoicing",
    ]


def main() -> int:
    ap = argparse.ArgumentParser(description="Allocate a new variation.")
    ap.add_argument("--project", required=True, help="Project code (e.g. 42SMITH)")
    ap.add_argument("--description", required=True, help="Short description (will be slugified)")
    ap.add_argument("--trigger", default="Client request",
                    help="One of the trigger types — see docs/email_parsing_checklist.md")
    ap.add_argument("--register", default=None,
                    help="Path to the project's master register. Defaults to "
                         "./working/<PROJECT>_Variation_Register.xlsx for local testing.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Don't write anything, just print the plan")
    args = ap.parse_args()

    if args.register:
        register_path = Path(args.register)
    else:
        LOCAL_WORKING.mkdir(exist_ok=True)
        register_path = LOCAL_WORKING / f"{args.project}_Variation_Register.xlsx"
        if not register_path.exists() and not args.dry_run:
            # Copy the template for local testing
            import shutil
            shutil.copy(TEMPLATE, register_path)
            # Stamp the project code into Project Setup B3
            wb = load_workbook(register_path)
            wb["Project Setup"]["B3"] = args.project
            wb.save(register_path)

    n = find_next_var_number(register_path) if register_path.exists() else 1
    var_id = f"{args.project}-VAR-{n:03d}"
    slug = slugify(args.description)
    dropbox_folder = f"/Renovate 8/Projects/{args.project} - .../03_Variations/{var_id} - {slug}"

    print(f"Variation ID:    {var_id}")
    print(f"Slug:            {slug}")
    print(f"Trigger:         {args.trigger}")
    print(f"Register:        {register_path}")
    print(f"Dropbox plan:")
    for p in planned_dropbox_paths(args.project, var_id, slug):
        print(f"   create folder: {p}")

    if args.dry_run:
        print("\n[dry-run] No files written.")
        return 0

    # Write the local register stub row
    row = stub_register_row(
        register_path,
        var_id=var_id,
        description=args.description,
        trigger=args.trigger,
        dropbox_folder=dropbox_folder,
    )
    print(f"\nStubbed register row {row} in {register_path} (status=RAISED, phase=1).")
    print("\nNext steps:")
    print("  1. Create the Dropbox folders listed above (use the Dropbox MCP).")
    print(f"  2. Save the original email + invoice into 01_Source/ in {var_id}.")
    var_suffix = var_id.split('-VAR-')[1]
    print(f"  3. Open templates/variation_notice_template.docx, save as a copy")
    print(f"     to 03_Notice/VAR-{var_suffix}_v1_RAISED.docx,")
    print( "     fill placeholders, and export to PDF.")
    print("  4. Email the PDF to the client.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
