"""Generate the per-project master Variation Register (Excel).

Run:
    python3 scripts/build_register_template.py

Writes:
    templates/master_register_template.xlsx

Layout
------
Sheet 1 — Variations      The line-by-line register, one row per variation.
Sheet 2 — Summary         Project totals, running contract sum.
Sheet 3 — Project Setup   Project-specific config (name, address, original sum, etc.).
Sheet 4 — Lookups         Status / Trigger / Phase enumerations for data validation.

Columns on Sheet 1 are grouped into:
    A–G   Identification, status, trigger, scope          (visible to anyone)
    H–N   COST TO ME (CONFIDENTIAL)                       (internal only)
    O–S   PRICE TO CLIENT (formula-driven from H–N)       (appears on notice)
    T–Z   Notice / approval / invoicing tracking          (workflow state)
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from pathlib import Path

OUTPUT = Path(__file__).resolve().parent.parent / "templates" / "master_register_template.xlsx"

# Colour palette
NAVY = "1F3864"
LIGHT_NAVY = "8FAADC"
CONF_RED = "C00000"
CONF_FILL = "FCE4E4"  # pale red for confidential columns
CLIENT_FILL = "E2EFDA"  # pale green for client-visible price columns
META_FILL = "FFF2CC"  # pale yellow for workflow / state columns
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
SUBHEADER_FONT = Font(name="Calibri", size=10, bold=True, italic=True, color=CONF_RED)
NORMAL_FONT = Font(name="Calibri", size=10)
CENTRE = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
THIN = Side(border_style="thin", color="808080")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def build():
    wb = Workbook()

    build_lookups_sheet(wb)
    build_setup_sheet(wb)
    build_variations_sheet(wb)
    build_summary_sheet(wb)

    # Remove the auto-created default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    # Reorder: Variations, Summary, Project Setup, Lookups
    wb.move_sheet("Variations", offset=-wb.sheetnames.index("Variations"))
    wb.active = 0

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT)
    print(f"Wrote {OUTPUT}")


def build_lookups_sheet(wb):
    ws = wb.create_sheet("Lookups")
    ws.sheet_state = "hidden"

    ws["A1"] = "Status"
    ws["B1"] = "Trigger"
    ws["C1"] = "Phase"
    ws["D1"] = "ApprovalMethod"
    ws["E1"] = "InvoiceMethod"
    for c in "ABCDE":
        ws[f"{c}1"].font = Font(bold=True)

    statuses = ["RAISED", "PRICED", "FINALISED", "APPROVED", "INVOICED", "CANCELLED", "ON HOLD"]
    triggers = [
        "Client request",
        "Site condition",
        "Council / certifier",
        "Architect / engineer",
        "Supplier-driven",
        "Builder-initiated",
        "Other",
    ]
    phases = ["1 — RAISED", "2 — FINALISED"]
    approval = ["Email confirmation", "Signed PDF", "Signed hardcopy", "Verbal + email confirm"]
    invoice = ["Next progress claim", "Standalone invoice", "Not yet invoiced"]

    for i, v in enumerate(statuses, start=2):
        ws.cell(row=i, column=1, value=v)
    for i, v in enumerate(triggers, start=2):
        ws.cell(row=i, column=2, value=v)
    for i, v in enumerate(phases, start=2):
        ws.cell(row=i, column=3, value=v)
    for i, v in enumerate(approval, start=2):
        ws.cell(row=i, column=4, value=v)
    for i, v in enumerate(invoice, start=2):
        ws.cell(row=i, column=5, value=v)

    # Define names for data validation
    wb.defined_names["StatusList"] = DefinedName("StatusList", attr_text=f"Lookups!$A$2:$A${1+len(statuses)}")
    wb.defined_names["TriggerList"] = DefinedName("TriggerList", attr_text=f"Lookups!$B$2:$B${1+len(triggers)}")
    wb.defined_names["PhaseList"] = DefinedName("PhaseList", attr_text=f"Lookups!$C$2:$C${1+len(phases)}")
    wb.defined_names["ApprovalList"] = DefinedName("ApprovalList", attr_text=f"Lookups!$D$2:$D${1+len(approval)}")
    wb.defined_names["InvoiceList"] = DefinedName("InvoiceList", attr_text=f"Lookups!$E$2:$E${1+len(invoice)}")


def build_setup_sheet(wb):
    ws = wb.create_sheet("Project Setup")
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 55

    title = ws["A1"]
    title.value = "PROJECT SETUP"
    title.font = Font(size=14, bold=True, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor=NAVY)
    title.alignment = CENTRE
    ws.merge_cells("A1:B1")
    ws.row_dimensions[1].height = 28

    rows = [
        ("Project code", "(e.g. 42SMITH)"),
        ("Project name / site address", ""),
        ("Client name(s)", ""),
        ("Client email", ""),
        ("Client phone", ""),
        ("Contract date", ""),
        ("Original contract sum (inc GST)", 0),
        ("Builder name", "Renovate 8"),
        ("Builder ABN", ""),
        ("Builder NSW Contractor License #", ""),
        ("Builder address", ""),
        ("Builder contact email", ""),
        ("Builder contact phone", ""),
        ("Contract form", "MBA NSW Residential Building Contract"),
        ("Dropbox root for project", ""),
        ("Default markup %", "(leave blank — set per variation)"),
    ]

    for i, (k, v) in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        ws.cell(row=i, column=1).fill = PatternFill("solid", fgColor=LIGHT_NAVY)
        ws.cell(row=i, column=1).alignment = LEFT
        ws.cell(row=i, column=1).border = BORDER
        cell = ws.cell(row=i, column=2, value=v)
        cell.alignment = LEFT
        cell.border = BORDER
        if k == "Original contract sum (inc GST)":
            cell.number_format = '"$"#,##0.00'

    # Define names so other sheets can reference these
    wb.defined_names["ProjectCode"] = DefinedName("ProjectCode", attr_text="'Project Setup'!$B$3")
    wb.defined_names["ProjectName"] = DefinedName("ProjectName", attr_text="'Project Setup'!$B$4")
    wb.defined_names["ClientName"] = DefinedName("ClientName", attr_text="'Project Setup'!$B$5")
    wb.defined_names["OriginalContractSum"] = DefinedName("OriginalContractSum", attr_text="'Project Setup'!$B$9")
    wb.defined_names["BuilderName"] = DefinedName("BuilderName", attr_text="'Project Setup'!$B$10")
    wb.defined_names["BuilderABN"] = DefinedName("BuilderABN", attr_text="'Project Setup'!$B$11")
    wb.defined_names["BuilderLicense"] = DefinedName("BuilderLicense", attr_text="'Project Setup'!$B$12")


def build_variations_sheet(wb):
    ws = wb.create_sheet("Variations")

    # Title row
    ws["A1"] = "VARIATION REGISTER"
    ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = CENTRE
    ws.merge_cells("A1:AC1")
    ws.row_dimensions[1].height = 26

    # Group header row (row 2)
    group_headers = [
        ("A", "G", "IDENTIFICATION & SCOPE", NAVY, "FFFFFF"),
        ("H", "N", "COST TO ME (CONFIDENTIAL — INTERNAL ONLY)", CONF_RED, "FFFFFF"),
        ("O", "S", "PRICE TO CLIENT (appears on notice)", "548235", "FFFFFF"),
        ("T", "AC", "NOTICE / APPROVAL / INVOICING", "BF8F00", "FFFFFF"),
    ]
    for start, end, label, fill, font in group_headers:
        ws.merge_cells(f"{start}2:{end}2")
        c = ws[f"{start}2"]
        c.value = label
        c.font = Font(bold=True, color=font, size=11)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = CENTRE
        c.border = BORDER
    ws.row_dimensions[2].height = 24

    # Column headers (row 3)
    cols = [
        ("A", "VAR #", 18, None, "text"),
        ("B", "Date Raised", 13, META_FILL, "date"),
        ("C", "Status", 14, META_FILL, "list:StatusList"),
        ("D", "Phase", 16, META_FILL, "list:PhaseList"),
        ("E", "Trigger", 22, META_FILL, "list:TriggerList"),
        ("F", "Short description", 35, None, "text"),
        ("G", "Detailed scope", 50, None, "text"),

        ("H", "Materials $", 13, CONF_FILL, "currency"),
        ("I", "Labour $", 13, CONF_FILL, "currency"),
        ("J", "Subcontractor $", 14, CONF_FILL, "currency"),
        ("K", "Plant / hire $", 13, CONF_FILL, "currency"),
        ("L", "Other $", 13, CONF_FILL, "currency"),
        ("M", "Cost to me (total)", 16, CONF_FILL, "formula"),
        ("N", "Markup %", 11, CONF_FILL, "percent"),

        ("O", "Price (ex GST)", 15, CLIENT_FILL, "formula"),
        ("P", "GST", 12, CLIENT_FILL, "formula"),
        ("Q", "Price (inc GST)", 16, CLIENT_FILL, "formula"),
        ("R", "Time impact (days)", 14, CLIENT_FILL, "number"),
        ("S", "Margin $", 13, CONF_FILL, "formula"),  # confidential: keep with cost group visually but logically tied to client price

        ("T", "Notice v1 issued", 14, META_FILL, "date"),
        ("U", "Notice v2 issued", 14, META_FILL, "date"),
        ("V", "Client approval date", 15, META_FILL, "date"),
        ("W", "Approval method", 18, META_FILL, "list:ApprovalList"),
        ("X", "Approved by (name)", 22, META_FILL, "text"),
        ("Y", "Invoice / claim #", 16, META_FILL, "text"),
        ("Z", "Invoice method", 18, META_FILL, "list:InvoiceList"),
        ("AA", "Date invoiced", 14, META_FILL, "date"),
        ("AB", "Dropbox folder", 40, None, "text"),
        ("AC", "Notes", 35, None, "text"),
    ]

    for col, label, width, fill, _kind in cols:
        cell = ws[f"{col}3"]
        cell.value = label
        cell.font = HEADER_FONT
        cell.alignment = CENTRE
        cell.border = BORDER
        if fill == CONF_FILL:
            cell.fill = PatternFill("solid", fgColor=CONF_RED)
        elif fill == CLIENT_FILL:
            cell.fill = PatternFill("solid", fgColor="548235")
        elif fill == META_FILL:
            cell.fill = PatternFill("solid", fgColor="BF8F00")
        else:
            cell.fill = PatternFill("solid", fgColor=NAVY)
        ws.column_dimensions[col].width = width
    ws.row_dimensions[3].height = 36

    # Add 50 data rows with formulas + formatting + validation
    first_data_row = 4
    last_data_row = 53

    # Data validations (each must be added once and applied to the column range)
    dv_map = {}
    for col, label, _w, _f, kind in cols:
        if isinstance(kind, str) and kind.startswith("list:"):
            list_name = kind.split(":", 1)[1]
            dv = DataValidation(type="list", formula1=f"={list_name}", allow_blank=True)
            dv.error = f"Pick from the {label} list"
            dv.errorTitle = "Invalid entry"
            dv.add(f"{col}{first_data_row}:{col}{last_data_row}")
            ws.add_data_validation(dv)
            dv_map[col] = dv

    # Row-by-row population
    for row in range(first_data_row, last_data_row + 1):
        for col, _label, _w, fill, kind in cols:
            cell = ws[f"{col}{row}"]
            cell.font = NORMAL_FONT
            cell.border = BORDER
            cell.alignment = LEFT if kind in ("text",) else CENTRE
            if fill == CONF_FILL:
                cell.fill = PatternFill("solid", fgColor=CONF_FILL)
            elif fill == CLIENT_FILL:
                cell.fill = PatternFill("solid", fgColor=CLIENT_FILL)
            elif fill == META_FILL:
                cell.fill = PatternFill("solid", fgColor=META_FILL)

            if kind == "currency":
                cell.number_format = '"$"#,##0.00;[Red]-"$"#,##0.00'
                cell.alignment = RIGHT
            elif kind == "percent":
                cell.number_format = "0.0%"
                cell.alignment = RIGHT
            elif kind == "number":
                cell.number_format = "0"
                cell.alignment = RIGHT
            elif kind == "date":
                cell.number_format = "yyyy-mm-dd"

        # Formulas
        # M (cost to me total) = SUM(H:L)
        ws[f"M{row}"] = f"=IFERROR(SUM(H{row}:L{row}),0)"
        ws[f"M{row}"].number_format = '"$"#,##0.00;[Red]-"$"#,##0.00'
        ws[f"M{row}"].alignment = RIGHT

        # O (Price ex GST) = M * (1 + N), but only if M and N have values
        ws[f"O{row}"] = f'=IF(AND(ISNUMBER(M{row}),ISNUMBER(N{row})),M{row}*(1+N{row}),"")'
        ws[f"O{row}"].number_format = '"$"#,##0.00;[Red]-"$"#,##0.00'
        ws[f"O{row}"].alignment = RIGHT

        # P (GST) = O * 10%
        ws[f"P{row}"] = f'=IF(ISNUMBER(O{row}),O{row}*0.1,"")'
        ws[f"P{row}"].number_format = '"$"#,##0.00;[Red]-"$"#,##0.00'
        ws[f"P{row}"].alignment = RIGHT

        # Q (inc GST) = O + P
        ws[f"Q{row}"] = f'=IF(AND(ISNUMBER(O{row}),ISNUMBER(P{row})),O{row}+P{row},"")'
        ws[f"Q{row}"].number_format = '"$"#,##0.00;[Red]-"$"#,##0.00'
        ws[f"Q{row}"].alignment = RIGHT

        # S (margin $) = O - M
        ws[f"S{row}"] = f'=IF(AND(ISNUMBER(O{row}),ISNUMBER(M{row})),O{row}-M{row},"")'
        ws[f"S{row}"].number_format = '"$"#,##0.00;[Red]-"$"#,##0.00'
        ws[f"S{row}"].alignment = RIGHT

    # Freeze top 3 rows + first column
    ws.freeze_panes = "B4"

    # Banner above the confidential block (row 2 is already coloured, add a one-cell warning above column H)
    # Actually let's add a footer note row instead
    note_row = last_data_row + 2
    ws[f"A{note_row}"] = ("CONFIDENTIAL columns (red header band) contain cost-to-me data and "
                          "MUST NEVER be shared with the client. CLIENT columns (green header band) "
                          "are the only price figures that appear on the formal Variation Notice.")
    ws[f"A{note_row}"].font = Font(italic=True, color=CONF_RED, size=10)
    ws[f"A{note_row}"].alignment = LEFT
    ws.merge_cells(f"A{note_row}:AC{note_row}")
    ws.row_dimensions[note_row].height = 30


def build_summary_sheet(wb):
    ws = wb.create_sheet("Summary")
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 60

    ws["A1"] = "PROJECT SUMMARY"
    ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = CENTRE
    ws.merge_cells("A1:C1")
    ws.row_dimensions[1].height = 26

    rows = [
        ("Project code", "=ProjectCode", ""),
        ("Project name / address", "=ProjectName", ""),
        ("Client", "=ClientName", ""),
        ("", "", ""),
        ("Original contract sum (inc GST)", "=OriginalContractSum", "From Project Setup"),
        ("Variations RAISED (count)", '=COUNTIF(Variations!C4:C53,"RAISED")', "Phase 1 issued"),
        ("Variations FINALISED+ (count)", '=COUNTIFS(Variations!C4:C53,"<>RAISED",Variations!C4:C53,"<>CANCELLED",Variations!C4:C53,"<>ON HOLD",Variations!C4:C53,"<>")', "Priced through to invoiced"),
        ("Variations CANCELLED (count)", '=COUNTIF(Variations!C4:C53,"CANCELLED")', ""),
        ("", "", ""),
        ("Total approved variations (inc GST)", '=IFERROR(SUMIFS(Variations!Q4:Q53,Variations!C4:C53,"APPROVED")+SUMIFS(Variations!Q4:Q53,Variations!C4:C53,"INVOICED"),0)', "Sum of APPROVED + INVOICED only"),
        ("Total pending variations (inc GST)", '=IFERROR(SUMIFS(Variations!Q4:Q53,Variations!C4:C53,"RAISED")+SUMIFS(Variations!Q4:Q53,Variations!C4:C53,"PRICED")+SUMIFS(Variations!Q4:Q53,Variations!C4:C53,"FINALISED"),0)', "Not yet client-approved"),
        ("", "", ""),
        ("Revised contract sum (approved only)", "=B6+B11", "Original + approved variations"),
        ("Revised contract sum (incl. pending)", "=B6+B11+B12", "Worst case — if everything pending also approves"),
        ("", "", ""),
        ("Total cost-to-me (CONFIDENTIAL)", '=IFERROR(SUMIFS(Variations!M4:M53,Variations!C4:C53,"APPROVED")+SUMIFS(Variations!M4:M53,Variations!C4:C53,"INVOICED"),0)', "Internal — DO NOT share"),
        ("Total margin (CONFIDENTIAL)", '=IFERROR(SUMIFS(Variations!S4:S53,Variations!C4:C53,"APPROVED")+SUMIFS(Variations!S4:S53,Variations!C4:C53,"INVOICED"),0)', "Internal — DO NOT share"),
        ("Average markup % (approved)", '=IFERROR(AVERAGEIFS(Variations!N4:N53,Variations!C4:C53,"APPROVED"),"")', "Internal"),
        ("Total time impact (days)", "=IFERROR(SUM(Variations!R4:R53),0)", "Across all non-cancelled rows"),
    ]

    for i, (k, v, note) in enumerate(rows, start=3):
        a = ws.cell(row=i, column=1, value=k)
        b = ws.cell(row=i, column=2, value=v)
        c = ws.cell(row=i, column=3, value=note)
        if k:
            a.font = Font(bold=True)
            a.fill = PatternFill("solid", fgColor=LIGHT_NAVY)
            a.alignment = LEFT
            a.border = BORDER
            b.alignment = RIGHT
            b.border = BORDER
            c.font = Font(italic=True, color="808080", size=9)
            c.alignment = LEFT
            c.border = BORDER
        # Currency formatting for $ rows
        if any(token in k for token in ("$", "sum", "Sum", "Total", "contract sum", "margin")) and "count" not in k.lower() and "days" not in k.lower() and "%" not in k:
            b.number_format = '"$"#,##0.00;[Red]-"$"#,##0.00'
        if "%" in k:
            b.number_format = "0.0%"
        if "(count)" in k:
            b.number_format = "0"
        if "days" in k.lower():
            b.number_format = "0"

        # Flag confidential rows
        if "CONFIDENTIAL" in k:
            a.font = Font(bold=True, color=CONF_RED)
            a.fill = PatternFill("solid", fgColor=CONF_FILL)
            b.fill = PatternFill("solid", fgColor=CONF_FILL)


if __name__ == "__main__":
    build()
