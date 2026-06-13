"""Populate the 10w variation register with all 15 Phase 2 variations."""
from __future__ import annotations

import datetime as dt
import shutil
from pathlib import Path

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = REPO_ROOT / "templates" / "master_register_template.xlsx"
WORKING_DIR = REPO_ROOT / "working"
OUT = WORKING_DIR / "10w.Variation Register.xlsx"

DROPBOX_BASE = "/Projects/Taren Point - 10 Woodlands Road/Variations/Phase 2"

VARIATIONS = [
    {
        "var": "V1", "date": "2026-02-19", "trigger": "Site condition",
        "desc": "Window BWIC",
        "scope": (
            "Building works in connection with window installation: (1) cutting concrete for "
            "master bedroom window frame; (2) Ardex/mastic joins to windows; "
            "(3) window protection; (4) crane hire."
        ),
        "folder": f"{DROPBOX_BASE}/V1 - Window BWIC",
    },
    {
        "var": "V2", "date": "2026-02-26", "trigger": "Client request",
        "desc": "Render",
        "scope": "Render works. Scope to be confirmed.",
        "folder": f"{DROPBOX_BASE}/V2 - Render",
    },
    {
        "var": "V3", "date": "2026-03-10", "trigger": "Site condition",
        "desc": "Lift BWIC",
        "scope": "Building works in connection with lift installation. Scope TBC.",
        "folder": f"{DROPBOX_BASE}/V3 - Lift BWIC",
    },
    {
        "var": "V4", "date": "2026-03-13", "trigger": "Site condition",
        "desc": "Garage Door Header",
        "scope": (
            "Garage door header modification: (1) install double LVL; "
            "(2) remove blue board wall and stud work; (3) reduce door width by 200mm."
        ),
        "folder": f"{DROPBOX_BASE}/V4 - Garage Door Header",
    },
    {
        "var": "V5", "date": "2026-04-02", "trigger": "Client request / Site condition",
        "desc": "Outside Works",
        "scope": (
            "Multiple external works: (1) side wheelchair ramp; "
            "(2) retaining wall at end of pool; (3) Mimi's side works; (4) rear yard works."
        ),
        "folder": f"{DROPBOX_BASE}/V5 - Outside Works",
    },
    {
        "var": "V6", "date": "2026-05-01", "trigger": "Client request",
        "desc": "Ground Floor — Built Out Wall to Living Room",
        "scope": "Construct built-out wall to living room, ground floor. Scope TBC.",
        "folder": f"{DROPBOX_BASE}/V6 - Ground Floor - Built out wall to Living Room",
    },
    {
        "var": "V7", "date": "2026-05-01", "trigger": "Client request",
        "desc": "Guest Bedroom Wall Nib Removal",
        "scope": "Remove wall nib in guest bedroom including all making good and reinstatement.",
        "folder": f"{DROPBOX_BASE}/V7 - Guest Bedroom Wall Nib Removal",
    },
    {
        "var": "V8", "date": "2026-05-01", "trigger": "Client request",
        "desc": "Rihanna Bedroom — Niche Installation",
        "scope": "Supply and install recessed niche in Rihanna's bedroom. Specification TBC.",
        "folder": f"{DROPBOX_BASE}/V8 - Rihanna Bedroom - Niche Installation",
    },
    {
        "var": "V9", "date": "2026-05-28", "trigger": "Site condition",
        "desc": "External Demolition — Mumu Side",
        "scope": (
            "External demolition works Mumu's side. "
            "Labour: 3 men × 1 day (28 May) + 1 man × 1 day (29 May). Price TBC."
        ),
        "folder": f"{DROPBOX_BASE}/V9 - External Demolition - Mumu Side",
    },
    {
        "var": "V10", "date": "2026-05-29", "trigger": "Client request",
        "desc": "Zip Tap Added to Level 1",
        "scope": "Supply and install Zip tap to Level 1 including all drainage lines. Plumber to price.",
        "folder": f"{DROPBOX_BASE}/V10 - Zip Tap Added to Level 1",
    },
    {
        "var": "V11", "date": "2026-05-31", "trigger": "Architect / engineer",
        "desc": "Additional Feature Dropped Ceilings",
        "scope": (
            "Additional feature dropped ceilings per amended RCP issued by architect 28 May 2026. "
            "Subcontractor quote outstanding."
        ),
        "folder": f"{DROPBOX_BASE}/V11 - Additional Feature Dropped Ceilings",
    },
    {
        "var": "V12", "date": "2026-06-06", "trigger": "Site condition",
        "desc": "Water Pipe to be Repaired by Plumber",
        "scope": "Water pipe repair by licensed plumber. Site condition. Plumber quote outstanding.",
        "folder": f"{DROPBOX_BASE}/V12 - Water Pipe to be Repaired by Plumber",
    },
    {
        "var": "V13", "date": "2026-06-01", "trigger": "Site condition",
        "desc": "Grouting Below Windows",
        "scope": (
            "Supply and apply Sika Grout GP and cement below windows. "
            "Materials cost: $463.20 inc GST (2 Euro Abrasives receipts). Labour TBC."
        ),
        "folder": f"{DROPBOX_BASE}/V13 - Grouting Below Windows",
    },
    {
        "var": "V14", "date": "2026-06-11", "trigger": "Client request",
        "desc": "Render — Smooth Finish to Existing Areas",
        "scope": (
            "Smooth render to all blue board areas (upgrade from scratch coat) + "
            "Rihanna's ramp side. Agreed on site 9 June 2026. Renderer quote TBC."
        ),
        "folder": f"{DROPBOX_BASE}/V14 - Render - Smooth Finish to Existing Areas",
    },
    {
        "var": "V15", "date": "2026-06-11", "trigger": "Client request",
        "desc": "External Waterproof Paint to Additional Areas",
        "scope": "External waterproof paint to additional areas. Scope and price TBC.",
        "folder": f"{DROPBOX_BASE}/V15 - External Waterproof Paint to Additional Areas",
    },
]

# Known materials costs (cost-to-me, inc GST) — confidential
MATERIALS = {
    "V5": 60.54,    # BC Sands cement (retaining wall area)
    "V13": 463.20,  # Euro Abrasives grout + cement (2 invoices)
}


def main():
    WORKING_DIR.mkdir(exist_ok=True)
    shutil.copy(TEMPLATE, OUT)

    wb = load_workbook(OUT)
    ws = wb["Variations"]

    # Find the "Project Setup" sheet and fill client email if possible
    # (skip if not present — the template handles it)

    for i, v in enumerate(VARIATIONS):
        row = 4 + i   # data starts at row 4 in the template

        date_obj = dt.date.fromisoformat(v["date"])

        ws.cell(row=row, column=1, value=v["var"])           # A: VAR #
        ws.cell(row=row, column=2, value=date_obj)           # B: Date Raised
        ws.cell(row=row, column=3, value="RAISED")           # C: Status
        ws.cell(row=row, column=4, value="1 — RAISED")       # D: Phase
        ws.cell(row=row, column=5, value=v["trigger"])       # E: Trigger
        ws.cell(row=row, column=6, value=v["desc"])          # F: Short description
        ws.cell(row=row, column=7, value=v["scope"])         # G: Detailed scope

        # Materials (col 8 = H) — confidential cost-to-me
        mat = MATERIALS.get(v["var"])
        if mat is not None:
            ws.cell(row=row, column=8, value=mat)

        # Notice v1 issued date — today (notices being issued now)
        ws.cell(row=row, column=20, value=dt.date.today())   # T: Notice v1 issued

        # Dropbox folder
        ws.cell(row=row, column=28, value=v["folder"])       # AB: Dropbox folder

        # Notes
        notice_note = f"v1 TBC notice issued {dt.date.today().strftime('%-d/%m/%Y')} — price outstanding"
        if v["var"] in MATERIALS:
            mat_str = f"${MATERIALS[v['var']]:,.2f}"
            notice_note += f"; materials cost-to-me {mat_str} inc GST on file"
        ws.cell(row=row, column=29, value=notice_note)       # AC: Notes

    wb.save(OUT)
    print(f"Written: {OUT}")
    print(f"  {len(VARIATIONS)} variation rows stubbed (all status=RAISED, phase=1—RAISED)")


if __name__ == "__main__":
    main()
