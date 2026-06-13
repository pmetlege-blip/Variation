"""Generate a summary Excel table for 10W Phase 2 variations.

Based on Dropbox listing + source document review done 2026-06-13.
Delivers a standalone summary (not the full register template).
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("pip install openpyxl")

OUT = Path(__file__).resolve().parent.parent / "working" / "10w_Phase2_Variations_Summary.xlsx"

NAVY_HEX  = "1F3864"
GREEN_HEX = "375623"
AMBER_HEX = "ED7D31"
RED_HEX   = "C00000"
LGREY_HEX = "F2F2F2"
DGREY_HEX = "D9D9D9"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

VARIATIONS = [
    {
        "var": "V1",
        "slug": "Window BWIC",
        "date_raised": "2026-02-19",
        "status": "RAISED",
        "trigger": "Site condition",
        "scope": (
            "Building works in connection with window installation: "
            "(1) cutting concrete for master bedroom window frame; "
            "(2) Ardex and mastic joins to windows; "
            "(3) window protection; "
            "(4) crane hire."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Non-standard (numbered sub-tasks)",
        "notes": "4 sub-task folders with site videos. No standard subfolders yet.",
    },
    {
        "var": "V2",
        "slug": "Render",
        "date_raised": "2026-02-26",
        "status": "RAISED",
        "trigger": "Client request",
        "scope": "Render works. Scope detail TBC — source material is a single site photo only.",
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": "Folder contains one JPG photo only. Scope needs confirming.",
    },
    {
        "var": "V3",
        "slug": "Lift BWIC",
        "date_raised": "2026-03-10",
        "status": "RAISED",
        "trigger": "Site condition",
        "scope": (
            "Building works in connection with lift installation. "
            "Site video footage only — detailed scope TBC."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": "2 MP4 site videos. No scope note on file.",
    },
    {
        "var": "V4",
        "slug": "Garage Door Header",
        "date_raised": "2026-03-13",
        "status": "RAISED",
        "trigger": "Site condition",
        "scope": (
            "Garage door to be flush with walls — modifications required to door header: "
            "(1) install double LVL; "
            "(2) remove blue board wall and stud work; "
            "(3) reduce garage door width by 200 mm. "
            "Additional cost to door supplier also expected."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": "Site meeting 13 Mar with Precision. Scope docx on file. Supplier quote outstanding.",
    },
    {
        "var": "V5",
        "slug": "Outside Works",
        "date_raised": "2026-04-02",
        "status": "RAISED",
        "trigger": "Client request / Site condition",
        "scope": (
            "Multiple external works: "
            "(1) side wheelchair ramp construction; "
            "(2) retaining wall required at end of pool; "
            "(3) Mimi's side external works; "
            "(4) rear yard works. "
            "Each sub-scope has separate site video/photo evidence."
        ),
        "materials_inc_gst": 60.54,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Non-standard (topical sub-folders)",
        "notes": (
            "Materials receipt: BC Sands $60.54 inc GST (cement, 10 Jun 26). "
            "Retaining wall quote/price outstanding. 4 sub-task folders with site footage."
        ),
    },
    {
        "var": "V6",
        "slug": "Ground Floor — Built Out Wall to Living Room",
        "date_raised": "2026-05-01",
        "status": "RAISED",
        "trigger": "Client request",
        "scope": (
            "Built-out wall to living room on ground floor. "
            "Source: WhatsApp screenshot. Detailed scope and price TBC."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": "WhatsApp screenshot only. Scope/price not yet documented.",
    },
    {
        "var": "V7",
        "slug": "Guest Bedroom Wall Nib Removal",
        "date_raised": "2026-05-01",
        "status": "RAISED",
        "trigger": "Client request",
        "scope": (
            "Removal of wall nib in guest bedroom. "
            "Site video footage on file. Detailed scope and price TBC."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": "9 MP4 site videos. No scope note on file.",
    },
    {
        "var": "V8",
        "slug": "Rihanna Bedroom — Niche Installation",
        "date_raised": "2026-05-01",
        "status": "RAISED",
        "trigger": "Client request",
        "scope": (
            "Installation of niche in Rihanna's bedroom. "
            "Site video footage on file. Detailed scope and price TBC."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": "9 MP4 site videos. No scope note on file.",
    },
    {
        "var": "V9",
        "slug": "External Demolition — Mumu Side",
        "date_raised": "2026-05-28",
        "status": "RAISED",
        "trigger": "Site condition",
        "scope": (
            "External demolition works on Mumu's side of property. "
            "Labour: 3 men × 1 day (28 May — Norm, Pierre, Boaz) "
            "+ 1 man × 1 day (29 May — Boaz)."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": (
            "Labour log on file (text note). 5 site videos. "
            "Price calculation pending — confirm day-rate to build cost."
        ),
    },
    {
        "var": "V10",
        "slug": "Zip Tap Added to Level 1",
        "date_raised": "2026-05-29",
        "status": "RAISED",
        "trigger": "Client request",
        "scope": (
            "Addition of Zip tap to Level 1, including drainage lines. "
            "Plumber (Chris) to provide price — awaiting quote."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": "Plumber quote outstanding. Site video on file.",
    },
    {
        "var": "V11",
        "slug": "Additional Feature Dropped Ceilings",
        "date_raised": "2026-05-31",
        "status": "RAISED",
        "trigger": "Architect / engineer",
        "scope": (
            "Additional feature dropped ceilings as per amended reflected ceiling plan "
            "issued by architect (amended RCP, 28 May 2026). "
            "Scope and price TBC pending subcontractor quote."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Non-standard (Received subfolder with architect's plan)",
        "notes": "Amended RCP PDF on file. Ceiling sub-contractor quote outstanding.",
    },
    {
        "var": "V12",
        "slug": "Water Pipe to be Repaired by Plumber",
        "date_raised": "2026-06-06",
        "status": "RAISED",
        "trigger": "Site condition",
        "scope": (
            "Water pipe discovered on site requiring repair by licensed plumber. "
            "Site photos on file. Plumber price TBC."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Partial (Photos sub-folder only)",
        "notes": "2 site photos on file. Plumber quote outstanding.",
    },
    {
        "var": "V13",
        "slug": "Grouting Below Windows",
        "date_raised": "2026-06-01",
        "status": "RAISED",
        "trigger": "Site condition",
        "scope": (
            "Grouting works below windows: supply and apply grout/cement as required. "
            "Materials purchased: Sika Grout GP + cement (Euro Abrasives, two purchases)."
        ),
        "materials_inc_gst": 463.20,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": (
            "Materials cost-to-me: $264.00 (Inv 10478145, 1 Jun) "
            "+ $199.20 (Inv 10478881, 10 Jun) = $463.20 inc GST. "
            "Labour component TBC. Client price TBC."
        ),
    },
    {
        "var": "V14",
        "slug": "Render — Smooth Finish to Existing Areas",
        "date_raised": "2026-06-11",
        "status": "RAISED",
        "trigger": "Client request",
        "scope": (
            "Smooth render finish (in lieu of scratch coat) to all blue board areas. "
            "Client agreed on site 9 June 2026: smooth finish throughout; "
            "Rihanna's ramp side also to receive smooth render (additional cost). "
            "Mum's side to retain scratch coat (no change)."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": "Client agreement documented in scope note. Renderer quote outstanding.",
    },
    {
        "var": "V15",
        "slug": "External Waterproof Paint to Additional Areas",
        "date_raised": "2026-06-11",
        "status": "RAISED",
        "trigger": "Client request",
        "scope": (
            "External waterproof paint applied to additional areas beyond original contract scope. "
            "Source: WhatsApp screenshot. Scope and price TBC."
        ),
        "materials_inc_gst": None,
        "labour_inc_gst": None,
        "sub_inc_gst": None,
        "price_inc_gst": None,
        "time_days": "TBC",
        "notice": "Not issued",
        "folder_structure": "Bare (no subfolders)",
        "notes": "WhatsApp screenshot only. Price/scope outstanding.",
    },
]

HEADERS = [
    "VAR #",
    "Date Raised",
    "Status",
    "Trigger",
    "Short Description",
    "Scope Summary",
    "Materials $ (cost-to-me, inc GST)",
    "Labour $ (cost-to-me, inc GST)",
    "Subcontractor $ (inc GST)",
    "PRICE TO CLIENT (inc GST)",
    "Time Impact (days)",
    "Notice Issued",
    "Folder Structure",
    "Notes / Outstanding Actions",
]

COL_WIDTHS = [8, 13, 10, 20, 30, 60, 18, 18, 18, 22, 15, 14, 25, 55]


def make_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "10W Phase 2 Variations"

    # Title row
    ws.merge_cells("A1:N1")
    title_cell = ws["A1"]
    title_cell.value = "10 WOODLANDS ROAD, TAREN POINT — PHASE 2 VARIATIONS SUMMARY"
    title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    title_cell.fill = fill(NAVY_HEX)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # Subtitle row
    ws.merge_cells("A2:N2")
    sub = ws["A2"]
    sub.value = f"Client: Chris and Joanna Gialouris  |  Builder: Renovate 8  |  As at: {dt.date.today().strftime('%-d %B %Y')}"
    sub.font = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
    sub.fill = fill("2E4A7A")
    sub.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    # Blank spacer
    ws.row_dimensions[3].height = 4

    # Header row
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        cell.fill = fill(NAVY_HEX)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()
    ws.row_dimensions[4].height = 28

    # Data rows
    for row_offset, v in enumerate(VARIATIONS):
        row = 5 + row_offset
        row_fill = fill(LGREY_HEX) if row_offset % 2 == 0 else fill("FFFFFF")

        def cell(col, value, num_format=None, bold=False, color=None, wrap=False, h_align="left"):
            c = ws.cell(row=row, column=col, value=value)
            c.font = Font(name="Calibri", size=9, bold=bold, color=color or "000000")
            c.fill = row_fill
            c.border = thin_border()
            c.alignment = Alignment(horizontal=h_align, vertical="top", wrap_text=wrap)
            if num_format:
                c.number_format = num_format
            return c

        cell(1, v["var"], bold=True, color=NAVY_HEX, h_align="center")
        d = dt.date.fromisoformat(v["date_raised"])
        cell(2, d, num_format="DD-MMM-YY", h_align="center")
        cell(3, v["status"], color=AMBER_HEX, bold=True, h_align="center")
        cell(4, v["trigger"])
        cell(5, v["slug"], bold=True, wrap=True)
        cell(6, v["scope"], wrap=True)

        # Cost columns — grey background to mark as CONFIDENTIAL
        conf_fill = fill("FFF2CC")
        for conf_col in (7, 8, 9):
            c = ws.cell(row=row, column=conf_col)
            c.fill = conf_fill
            c.border = thin_border()
            c.font = Font(name="Calibri", size=9, color="7F7F7F", italic=True)
            c.alignment = Alignment(horizontal="right", vertical="top")

        mat = v["materials_inc_gst"]
        lab = v["labour_inc_gst"]
        sub = v["sub_inc_gst"]
        ws.cell(row=row, column=7).value = mat if mat is not None else "TBC"
        ws.cell(row=row, column=8).value = lab if lab is not None else "TBC"
        ws.cell(row=row, column=9).value = sub if sub is not None else "TBC"
        if isinstance(mat, (int, float)):
            ws.cell(row=row, column=7).number_format = '"$"#,##0.00'
        if isinstance(lab, (int, float)):
            ws.cell(row=row, column=8).number_format = '"$"#,##0.00'
        if isinstance(sub, (int, float)):
            ws.cell(row=row, column=9).number_format = '"$"#,##0.00'

        price = v["price_inc_gst"]
        price_cell = cell(10, price if price is not None else "TBC",
                          num_format='"$"#,##0.00' if isinstance(price, float) else None,
                          bold=isinstance(price, float), h_align="right")

        cell(11, v["time_days"], h_align="center")
        notice_color = "7F7F7F"
        cell(12, v["notice"], color=notice_color, h_align="center")
        cell(13, v["folder_structure"], wrap=True)
        cell(14, v["notes"], wrap=True)

        ws.row_dimensions[row].height = 48

    # Confidence note below data
    note_row = 5 + len(VARIATIONS) + 1
    ws.merge_cells(f"A{note_row}:N{note_row}")
    note = ws[f"A{note_row}"]
    note.value = (
        "CONFIDENTIAL NOTE: Columns G–I (yellow) contain cost-to-me data only. "
        "Do NOT include on client-facing documents. All variations are currently Phase 1 — RAISED; "
        "no formal Variation Notices have yet been issued for this project."
    )
    note.font = Font(name="Calibri", size=8, italic=True, color="7F7F7F")
    note.fill = fill("FFF2CC")
    note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[note_row].height = 24

    # Column widths
    for col_idx, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Freeze panes
    ws.freeze_panes = "B5"

    return wb


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = make_workbook()
    wb.save(OUT)
    print(f"Written: {OUT}")


if __name__ == "__main__":
    main()
