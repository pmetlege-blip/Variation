"""Allocate a new variation in a Renovate 8 project.

What it does
------------
- Reads the next variation number from the project's existing
  /Projects/<Suburb> - <Address>/Variations/ folder. Numbering is V# —
  per-project sequential, matching the existing Renovate 8 convention.
- Prints the planned Dropbox folder tree (5 standard subfolders inside
  the new V# - <slug> folder).
- Optionally stubs a row in the project's master register
  (Variation Register.xlsx) at status RAISED.

Folders are NOT created by this script directly — Claude creates them
via the Dropbox MCP tools after running this script, so the user can
confirm the plan first.

Usage:
    python3 scripts/create_variation.py \\
        --project-folder "Annandale - 137 Annandale Street" \\
        --description "extra power points kitchen" \\
        --trigger "Client request"

    # Dry run (no register stub):
    python3 scripts/create_variation.py --project-folder "..." \\
        --description "..." --dry-run

    # Force a specific next number (if migrating mid-project):
    python3 scripts/create_variation.py --project-folder "..." \\
        --description "..." --number 7
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
CONFIG = REPO_ROOT / "scripts" / "config.json"
LOCAL_WORKING = REPO_ROOT / "working"

# Standard per-variation subfolders (created INSIDE each V# folder)
PER_VARIATION_SUBFOLDERS = [
    "01_Source",
    "01_Source/photos",
    "02_Costing",
    "03_Notice",
    "04_Approval",
    "05_Invoicing",
]


def load_config() -> dict:
    if not CONFIG.exists():
        return {}
    with open(CONFIG) as f:
        return json.load(f)


def slugify(s: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", " ", s).strip()
    return s[:60]


def find_next_v_number(register_path: Path | None,
                       existing_v_numbers: list[int] | None = None) -> int:
    """Return the next V# (1-indexed) for the project.

    Two sources of truth, in order:
    1. Numbers passed in via `existing_v_numbers` (from a Dropbox listing).
    2. VAR # column of the project's register if it exists.
    """
    used: list[int] = []
    if existing_v_numbers:
        used.extend(existing_v_numbers)
    if register_path and register_path.exists():
        wb = load_workbook(register_path)
        ws = wb["Variations"]
        for row in ws.iter_rows(min_row=4, max_row=ws.max_row, max_col=1, values_only=True):
            val = row[0]
            if not val:
                continue
            m = re.match(r"V(\d+)\b", str(val))
            if m:
                used.append(int(m.group(1)))
    return (max(used) + 1) if used else 1


def stub_register_row(register_path: Path, *, var_id: str, description: str,
                      trigger: str, dropbox_folder: str) -> int:
    """Append a new RAISED row. Returns the row number used."""
    wb = load_workbook(register_path)
    ws = wb["Variations"]
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


def planned_dropbox_paths(project_folder: str, var_id: str, slug: str,
                          *, dropbox_root: str = "/Projects",
                          variations_subfolder: str = "Variations") -> list[str]:
    folder_name = f"{var_id} - {slug}" if slug else var_id
    base = f"{dropbox_root}/{project_folder}/{variations_subfolder}/{folder_name}"
    return [base] + [f"{base}/{sub}" for sub in PER_VARIATION_SUBFOLDERS]


def main() -> int:
    cfg = load_config()
    dbx = cfg.get("dropbox", {})
    dropbox_root = dbx.get("root", "/Projects")
    variations_subfolder = dbx.get("variations_subfolder", "Variations")
    register_filename = dbx.get("register_filename", "Variation Register.xlsx")

    ap = argparse.ArgumentParser(description="Allocate a new V# variation.")
    ap.add_argument("--project-folder", required=True,
                    help='Dropbox folder name under /Projects, '
                         'e.g. "Annandale - 137 Annandale Street"')
    ap.add_argument("--description", required=True,
                    help="Short description (will be cleaned for the folder name)")
    ap.add_argument("--trigger", default="Client request",
                    help="One of the trigger types — see docs/email_parsing_checklist.md")
    ap.add_argument("--number", type=int, default=None,
                    help="Force a specific V number (skips auto-allocate)")
    ap.add_argument("--existing", default="",
                    help="Comma-separated list of existing V numbers in the Dropbox "
                         "Variations folder (e.g. \"1,2,3,4\"). Passed in by Claude "
                         "after listing the folder via Dropbox MCP.")
    ap.add_argument("--register", default=None,
                    help="Path to the project's master register. Defaults to "
                         "./working/<folder>/<register_filename> for local testing.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Don't write anything, just print the plan")
    args = ap.parse_args()

    project_folder = args.project_folder

    # Resolve register path. Locally, working/<sanitized_folder>/<filename>.
    if args.register:
        register_path = Path(args.register)
    else:
        LOCAL_WORKING.mkdir(exist_ok=True)
        safe_folder = re.sub(r"[^A-Za-z0-9._ -]", "_", project_folder)
        register_dir = LOCAL_WORKING / safe_folder
        register_dir.mkdir(parents=True, exist_ok=True)
        register_path = register_dir / register_filename
        if not register_path.exists() and not args.dry_run:
            import shutil
            shutil.copy(TEMPLATE, register_path)

    existing_v = []
    if args.existing.strip():
        existing_v = [int(x) for x in re.findall(r"\d+", args.existing)]

    if args.number is not None:
        n = args.number
    else:
        n = find_next_v_number(register_path if register_path.exists() else None,
                               existing_v_numbers=existing_v)

    var_id = f"V{n}"
    slug = slugify(args.description)
    folder_name = f"{var_id} - {slug}" if slug else var_id
    dropbox_folder = f"{dropbox_root}/{project_folder}/{variations_subfolder}/{folder_name}"

    print(f"Project folder:   {project_folder}")
    print(f"Variation ID:     {var_id}")
    print(f"Slug:             {slug}")
    print(f"Trigger:          {args.trigger}")
    print(f"Local register:   {register_path}")
    print(f"Dropbox plan:")
    for p in planned_dropbox_paths(project_folder, var_id, slug,
                                    dropbox_root=dropbox_root,
                                    variations_subfolder=variations_subfolder):
        print(f"   create folder: {p}")

    if args.dry_run:
        print("\n[dry-run] No files written.")
        return 0

    row = stub_register_row(
        register_path,
        var_id=var_id,
        description=args.description,
        trigger=args.trigger,
        dropbox_folder=dropbox_folder,
    )
    print(f"\nStubbed register row {row} in {register_path} (status=RAISED, phase=1).")
    print("\nNext steps (Claude executes via Dropbox MCP):")
    print(f"  1. dropbox.create_folder for each path above.")
    print(f"  2. Save the source email + invoice into {folder_name}/01_Source/.")
    print(f"  3. Generate Notice from templates/variation_notice_template.docx,")
    print(f"     save as 03_Notice/{var_id}_v1_RAISED.docx, export to PDF.")
    print(f"  4. Upload the updated register back to Dropbox at:")
    print(f"     {dropbox_root}/{project_folder}/{variations_subfolder}/{register_filename}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
