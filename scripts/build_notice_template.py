"""Generate the Variation Notice template (Word).

Run:
    python3 scripts/build_notice_template.py

Writes:
    templates/variation_notice_template.docx

The template uses {{double-curly}} placeholders so it can be filled
programmatically (later script) or by hand. It supports both PHASE 1
(RAISED) and PHASE 2 (FINALISED) — toggle by editing the status banner
and the price section preamble.

Compliance reference: docs/legal_compliance.md
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path
import json

REPO_ROOT = Path(__file__).resolve().parent.parent
OUTPUT = REPO_ROOT / "templates" / "variation_notice_template.docx"
CONFIG = REPO_ROOT / "scripts" / "config.json"


def load_builder_fields():
    """Read scripts/config.json if present and return a dict of builder placeholder replacements.
    Falls back to keeping {{BUILDER_*}} placeholders if config is absent."""
    fields = {
        "BUILDER_ADDRESS": "{{BUILDER_ADDRESS}}",
        "BUILDER_ABN": "{{BUILDER_ABN}}",
        "BUILDER_LICENCE": "{{BUILDER_LICENCE}}",
        "BUILDER_EMAIL": "{{BUILDER_EMAIL}}",
        "BUILDER_PHONE": "{{BUILDER_PHONE}}",
        "BUILDER_SIGNATORY": "{{BUILDER_SIGNATORY}}",
    }
    if CONFIG.exists():
        with open(CONFIG) as f:
            cfg = json.load(f)
        b = cfg.get("builder", {})
        if b.get("address"): fields["BUILDER_ADDRESS"] = b["address"]
        if b.get("abn"): fields["BUILDER_ABN"] = b["abn"]
        if b.get("licence"): fields["BUILDER_LICENCE"] = b["licence"]
        if b.get("email"): fields["BUILDER_EMAIL"] = b["email"]
        if b.get("phone"): fields["BUILDER_PHONE"] = b["phone"]
        if b.get("signatory"): fields["BUILDER_SIGNATORY"] = b["signatory"]
    return fields


BUILDER = load_builder_fields()

NAVY = RGBColor(0x1F, 0x38, 0x64)
RED = RGBColor(0xC0, 0x00, 0x00)
GREY = RGBColor(0x59, 0x59, 0x59)


def shade(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    if level == 0:
        run.font.size = Pt(18)
        run.font.color.rgb = NAVY
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif level == 1:
        run.font.size = Pt(13)
        run.font.color.rgb = NAVY
    else:
        run.font.size = Pt(11)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    return p


def add_para(doc, text, *, italic=False, bold=False, color=None, size=10, align=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.italic = italic
    run.bold = bold
    run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(4)
    return p


def add_field_row(table, label, value_placeholder, *, label_w_cm=5.5):
    row = table.add_row()
    row.cells[0].text = ""
    row.cells[1].text = ""
    p0 = row.cells[0].paragraphs[0]
    r0 = p0.add_run(label)
    r0.bold = True
    r0.font.size = Pt(10)
    p1 = row.cells[1].paragraphs[0]
    r1 = p1.add_run(value_placeholder)
    r1.font.size = Pt(10)
    row.cells[0].width = Cm(label_w_cm)
    shade(row.cells[0], "E7E6E6")


def build():
    doc = Document()

    # Page setup
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    # === Header band ===
    # If a logo image exists at templates/logo.{png,jpg,jpeg}, insert it in
    # place of the "RENOVATE 8" text heading. Otherwise fall back to text.
    logo_candidates = [
        REPO_ROOT / "templates" / "logo.png",
        REPO_ROOT / "templates" / "logo.jpg",
        REPO_ROOT / "templates" / "logo.jpeg",
    ]
    logo_path = next((p for p in logo_candidates if p.exists()), None)
    if logo_path is not None:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(str(logo_path), width=Cm(6.0))
    else:
        add_heading(doc, "RENOVATE 8", level=0)
    add_para(
        doc,
        f'{BUILDER["BUILDER_ADDRESS"]}    |    ABN {BUILDER["BUILDER_ABN"]}    |    NSW Contractor Licence {BUILDER["BUILDER_LICENCE"]}',
        italic=True, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, size=9,
    )
    add_para(
        doc,
        f'{BUILDER["BUILDER_EMAIL"]}    |    {BUILDER["BUILDER_PHONE"]}',
        italic=True, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, size=9,
    )

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    run = p.add_run("FORMAL VARIATION NOTICE")
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = NAVY

    # Status banner
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("{{STATUS_BANNER}}")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RED
    add_para(
        doc,
        'Replace {{STATUS_BANNER}} with one of:  "PHASE 1 — RAISED (PRICE TO BE CONFIRMED)"  •  '
        '"PHASE 1 — RAISED (PRICE CONFIRMED — AWAITING ACKNOWLEDGMENT)"  •  '
        '"PHASE 2 — FINALISED (FINAL PRICE)"',
        italic=True, color=GREY, size=8, align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    add_para(
        doc,
        "Issued under the Master Builders Association of NSW Residential Building Contract "
        "and section 7AAA of the Home Building Act 1989 (NSW).",
        italic=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY,
    )

    # === Identification block ===
    t = doc.add_table(rows=0, cols=2)
    t.autofit = False
    t.columns[0].width = Cm(5.5)
    t.columns[1].width = Cm(11.5)
    add_field_row(t, "Project / site address", "{{PROJECT_ADDRESS}}")
    add_field_row(t, "Client (owner)", "{{CLIENT_NAME}}")
    add_field_row(t, "Head contract dated", "{{CONTRACT_DATE}}")
    add_field_row(t, "Variation number", "{{VAR_NUMBER}}")
    add_field_row(t, "Notice version", "{{NOTICE_VERSION}}")
    add_field_row(t, "Date issued", "{{DATE_ISSUED}}")
    add_field_row(t, "Date variation identified", "{{DATE_TRIGGER}}")
    add_field_row(t, "Trigger / source", "{{TRIGGER}}")

    # === 1. Nature of variation ===
    add_heading(doc, "1. Nature of variation", level=1)
    add_para(doc, "{{DESCRIPTION_SHORT}}")

    # === 2. Reason ===
    add_heading(doc, "2. Reason for variation", level=1)
    add_para(doc, "{{REASON}}")

    # === 3. Scope ===
    add_heading(doc, "3. Scope of works", level=1)
    add_para(
        doc,
        "Inclusions: {{SCOPE_INCLUSIONS}}",
    )
    add_para(
        doc,
        "Exclusions: {{SCOPE_EXCLUSIONS}}",
    )
    add_para(
        doc,
        "Drawing / specification reference: {{DRAWING_REF}}",
    )

    # === 4. Price ===
    add_heading(doc, "4. Price impact", level=1)
    add_para(
        doc,
        "This variation will {{PRICE_DIRECTION}} the contract sum by the following amount. "
        "All figures shown are the price payable by the Client. Where this notice is issued "
        "as Phase 1 with an indicative estimate, the figures below are subject to confirmation "
        "and a Phase 2 (Finalised) notice will be reissued.",
    )

    pricet = doc.add_table(rows=1, cols=2)
    pricet.autofit = False
    pricet.columns[0].width = Cm(8.5)
    pricet.columns[1].width = Cm(5.5)

    hdr_cells = pricet.rows[0].cells
    hdr_cells[0].text = ""
    hdr_cells[1].text = ""
    r = hdr_cells[0].paragraphs[0].add_run("Item")
    r.bold = True
    r = hdr_cells[1].paragraphs[0].add_run("Amount (AUD)")
    r.bold = True
    hdr_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    shade(hdr_cells[0], "1F3864")
    shade(hdr_cells[1], "1F3864")
    for c in hdr_cells:
        for p in c.paragraphs:
            for run in p.runs:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    def price_row(label, placeholder, *, bold=False, fill=None):
        r = pricet.add_row()
        r.cells[0].text = ""
        r.cells[1].text = ""
        p0 = r.cells[0].paragraphs[0]
        run0 = p0.add_run(label)
        run0.bold = bold
        run0.font.size = Pt(10)
        p1 = r.cells[1].paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run1 = p1.add_run(placeholder)
        run1.bold = bold
        run1.font.size = Pt(10)
        if fill:
            shade(r.cells[0], fill)
            shade(r.cells[1], fill)

    price_row("Sub-total (ex GST)", "{{PRICE_EX_GST}}")
    price_row("GST (10%)", "{{PRICE_GST}}")
    price_row("TOTAL (inc GST)", "{{PRICE_INC_GST}}", bold=True, fill="E2EFDA")

    add_para(
        doc,
        "Pricing basis: {{PRICING_BASIS}}",
        italic=True, size=9, color=GREY,
    )
    add_para(
        doc,
        "(Pricing basis examples — operator picks one and rewrites in plain language:  "
        "\"Fixed price quoted by [supplier] dated [date]\";  "
        "\"Per supplier invoice [#] plus builder labour and overhead\";  "
        "\"Indicative — to be confirmed once subcontractor quote received\". "
        "Prefer fixed-price wording. Do not state an internal markup percentage on this notice.)",
        italic=True, size=8, color=GREY,
    )

    # === 5. Time impact ===
    add_heading(doc, "5. Time impact", level=1)
    add_para(
        doc,
        "This variation will {{TIME_DIRECTION}} the construction period by {{TIME_DAYS}} working day(s).",
    )
    add_para(
        doc,
        "Revised practical completion date (subject to weather and other contractual extensions): "
        "{{REVISED_PC_DATE}}.",
    )

    # === 6. Effect on contract sum ===
    add_heading(doc, "6. Effect on contract sum", level=1)
    cs = doc.add_table(rows=1, cols=2)
    cs.autofit = False
    cs.columns[0].width = Cm(10.5)
    cs.columns[1].width = Cm(5.5)
    cs.rows[0].cells[0].text = ""
    cs.rows[0].cells[1].text = ""
    r = cs.rows[0].cells[0].paragraphs[0].add_run("Item"); r.bold = True
    r = cs.rows[0].cells[1].paragraphs[0].add_run("Amount (inc GST, AUD)"); r.bold = True
    cs.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    shade(cs.rows[0].cells[0], "1F3864")
    shade(cs.rows[0].cells[1], "1F3864")
    for c in cs.rows[0].cells:
        for p in c.paragraphs:
            for run in p.runs:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    def sum_row(label, placeholder, *, bold=False, fill=None):
        r = cs.add_row()
        r.cells[0].text = ""
        r.cells[1].text = ""
        p0 = r.cells[0].paragraphs[0]
        run0 = p0.add_run(label); run0.bold = bold; run0.font.size = Pt(10)
        p1 = r.cells[1].paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run1 = p1.add_run(placeholder); run1.bold = bold; run1.font.size = Pt(10)
        if fill:
            shade(r.cells[0], fill)
            shade(r.cells[1], fill)

    sum_row("Original contract sum", "{{CS_ORIGINAL}}")
    sum_row("Plus: previously approved variations", "{{CS_PREVIOUS_VARS}}")
    sum_row("Plus / minus: this variation", "{{CS_THIS_VAR}}")
    sum_row("REVISED CONTRACT SUM", "{{CS_NEW_TOTAL}}", bold=True, fill="E2EFDA")

    # === 7. Payment ===
    add_heading(doc, "7. Payment", level=1)
    add_para(
        doc,
        "This variation will be invoiced as follows (one option applies):",
    )
    add_para(
        doc,
        "[  ] Bundled into the next progress claim (preferred, in accordance with the head contract).",
    )
    add_para(
        doc,
        "[  ] By a separate tax invoice issued on completion of this variation.",
    )
    add_para(
        doc,
        "Payment terms: {{PAYMENT_TERMS}}",
    )

    # === 8. Statutory warranties ===
    add_heading(doc, "8. Statutory warranties", level=1)
    add_para(
        doc,
        "All work the subject of this variation is performed subject to the statutory warranties "
        "set out in Schedule 2 of the Home Building Act 1989 (NSW), which apply notwithstanding "
        "anything in this notice.",
        size=9,
    )

    # === 9. Emergency / urgency carve-out ===
    add_heading(doc, "9. Urgency / commencement", level=1)
    add_para(
        doc,
        "[  ] Standard — Work on this variation will not commence until this notice has been "
        "signed by the Client (clause 10 below).",
        size=10,
    )
    add_para(
        doc,
        "[  ] Urgent / safety — Work has commenced (or will commence) before signature because it is "
        "necessary for safety, compliance with law, or to prevent further loss or damage. The Client "
        "is asked to sign this notice promptly. Reason for urgency: {{URGENCY_REASON}}.",
        size=10,
    )

    # === 10. Client acknowledgment ===
    add_heading(doc, "10. Client acknowledgment & acceptance", level=1)
    add_para(
        doc,
        "By signing below, the Client acknowledges that:",
    )
    for bullet in [
        "the variation has been explained to them;",
        "they consent to the variation being performed;",
        "the price and time impact set out above are accepted;",
        "they have had the opportunity to ask questions and (where they wished) seek independent advice;",
        "this acceptance forms part of the head contract and is a variation under section 7AAA of the Home Building Act 1989 (NSW).",
    ]:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(bullet)
        run.font.size = Pt(10)

    sigt = doc.add_table(rows=2, cols=2)
    sigt.autofit = False
    sigt.columns[0].width = Cm(8.5)
    sigt.columns[1].width = Cm(8.5)

    def sig_cell(cell, label, name_field, date_field):
        cell.text = ""
        p1 = cell.paragraphs[0]
        r = p1.add_run(label); r.bold = True; r.font.size = Pt(10)
        p2 = cell.add_paragraph(); p2.paragraph_format.space_before = Pt(18)
        r = p2.add_run("Signature: ______________________________________"); r.font.size = Pt(10)
        p3 = cell.add_paragraph()
        r = p3.add_run(f"Print name: {name_field}"); r.font.size = Pt(10)
        p4 = cell.add_paragraph()
        r = p4.add_run(f"Date: {date_field}"); r.font.size = Pt(10)

    sig_cell(sigt.rows[0].cells[0], "CLIENT (Owner)", "{{CLIENT_NAME}}", "________________")
    sig_cell(sigt.rows[0].cells[1], "CLIENT (Co-owner, if applicable)", "{{CLIENT_NAME_2}}", "________________")
    sig_cell(sigt.rows[1].cells[0], "RENOVATE 8 (Builder)", BUILDER["BUILDER_SIGNATORY"], "{{DATE_ISSUED}}")
    sig_cell(sigt.rows[1].cells[1], "Witness (if required)", "________________", "________________")

    # === Footer notice ===
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(
        "This is a formal variation notice. Please return a signed copy to Renovate 8 "
        "(by email reply with signed PDF, or by hand)."
    )
    r.italic = True
    r.font.color.rgb = GREY
    r.font.size = Pt(9)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(
        f'Renovate 8  •  ABN {BUILDER["BUILDER_ABN"]}  •  NSW Contractor Licence {BUILDER["BUILDER_LICENCE"]}  •  '
        "Filed: {{VAR_NUMBER}} / v{{NOTICE_VERSION}}"
    )
    r.italic = True
    r.font.color.rgb = GREY
    r.font.size = Pt(8)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    build()
