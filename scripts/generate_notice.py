"""Generate a FILLED Variation Notice (Word .docx, optional PDF).

Unlike build_notice_template.py (which writes the blank {{placeholder}}
template), this script produces a ready-to-send Notice with real values, and
**switches the status banner and the price section automatically** based on
the variation's price posture (--shape):

    tbc        Phase 1 — price genuinely unknown ("blank cheque"). No price
               table; a statement that no amount is payable yet.
    indicative Phase 1 — an estimate or range, clearly marked "estimate only".
    fixed      Phase 1 — firm price, awaiting the client's signature.
    finalised  Phase 2 — final price, the version the client signs.

Contract identity (client, project address, contract date, original sum) is
read from scripts/config.json's projects registry by --project-code, so you
only pass the variation-specific details.

Examples
--------
# Day 1, site finding, no price yet:
python3 scripts/generate_notice.py --project-code 137a --var V5 --shape tbc \
    --description "Rotten subfloor joists found under main bathroom" \
    --reason "Concealed rot exposed when floor tiles were lifted" \
    --scope-inclusions "Remove affected joists/bearers; supply and install new treated timber; reinstate floor substrate" \
    --trigger "Site condition" --date-identified 2026-05-15

# A week later, a rough range:
python3 scripts/generate_notice.py --project-code 137a --var V5 --version 2 \
    --shape indicative --price-low 4000 --price-high 6000 \
    --description "..." --trigger "Site condition" --date-identified 2026-05-15

# Firmed up, fixed price awaiting signature:
python3 scripts/generate_notice.py --project-code 137a --var V5 --version 3 \
    --shape fixed --price-inc-gst 6247.50 --time-days 2 --description "..."

# Final, client signs this one:
python3 scripts/generate_notice.py --project-code 137a --var V5 --version 4 \
    --shape finalised --price-inc-gst 6247.50 --time-days 2 \
    --previous-approved 18500 --description "..." --pdf
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

REPO_ROOT = Path(__file__).resolve().parent.parent
CONFIG = REPO_ROOT / "scripts" / "config.json"
WORKING = REPO_ROOT / "working"

NAVY = RGBColor(0x1F, 0x38, 0x64)
RED = RGBColor(0xC0, 0x00, 0x00)
GREEN = RGBColor(0x37, 0x7D, 0x22)
GREY = RGBColor(0x59, 0x59, 0x59)

# Banner text + filing state per shape. RAISED = Phase 1, FINALISED = Phase 2.
SHAPES = {
    "tbc": {
        "banner": "PHASE 1 — RAISED (PRICE TO BE CONFIRMED)",
        "state": "RAISED",
        "banner_color": RED,
    },
    "indicative": {
        "banner": "PHASE 1 — RAISED (INDICATIVE PRICE — ESTIMATE ONLY)",
        "state": "RAISED",
        "banner_color": RED,
    },
    "fixed": {
        "banner": "PHASE 1 — RAISED (PRICE CONFIRMED — AWAITING ACKNOWLEDGMENT)",
        "state": "RAISED",
        "banner_color": RED,
    },
    "finalised": {
        "banner": "PHASE 2 — FINALISED (FINAL PRICE)",
        "state": "FINALISED",
        "banner_color": GREEN,
    },
}


# ---------------------------------------------------------------- config

def load_config() -> dict:
    if not CONFIG.exists():
        return {}
    with open(CONFIG) as f:
        return json.load(f)


def money(x) -> str:
    return f"${x:,.2f}"


def fmt_date(s: str) -> str:
    """Render an ISO date as e.g. 20 August 2025; pass through anything else."""
    if not s:
        return ""
    try:
        return dt.date.fromisoformat(str(s)).strftime("%-d %B %Y")
    except (ValueError, TypeError):
        return str(s)


# ---------------------------------------------------------------- styling

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
        run.font.size = Pt(18); run.font.color.rgb = NAVY
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif level == 1:
        run.font.size = Pt(13); run.font.color.rgb = NAVY
    else:
        run.font.size = Pt(11)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    return p


def add_para(doc, text, *, italic=False, bold=False, color=None, size=10, align=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.italic = italic; run.bold = bold; run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(4)
    return p


def add_field_row(table, label, value):
    row = table.add_row()
    row.cells[0].text = ""; row.cells[1].text = ""
    r0 = row.cells[0].paragraphs[0].add_run(label); r0.bold = True; r0.font.size = Pt(10)
    r1 = row.cells[1].paragraphs[0].add_run(value); r1.font.size = Pt(10)
    row.cells[0].width = Cm(5.5)
    shade(row.cells[0], "E7E6E6")


def money_table(doc, rows):
    """rows: list of (label, value, bold, fill_or_None)."""
    t = doc.add_table(rows=1, cols=2)
    t.autofit = False
    t.columns[0].width = Cm(10.5)
    t.columns[1].width = Cm(5.5)
    h = t.rows[0].cells
    h[0].text = ""; h[1].text = ""
    r = h[0].paragraphs[0].add_run("Item"); r.bold = True
    r = h[1].paragraphs[0].add_run("Amount (AUD)"); r.bold = True
    h[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    shade(h[0], "1F3864"); shade(h[1], "1F3864")
    for c in h:
        for pp in c.paragraphs:
            for run in pp.runs:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for label, value, bold, fill in rows:
        row = t.add_row()
        row.cells[0].text = ""; row.cells[1].text = ""
        r0 = row.cells[0].paragraphs[0].add_run(label); r0.bold = bold; r0.font.size = Pt(10)
        p1 = row.cells[1].paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r1 = p1.add_run(value); r1.bold = bold; r1.font.size = Pt(10)
        if fill:
            shade(row.cells[0], fill); shade(row.cells[1], fill)
    return t


# ---------------------------------------------------------------- pricing

def compute_prices(args, gst_rate):
    """Return a dict describing the price posture for the chosen shape."""
    shape = args.shape

    def from_inc(inc):
        ex = inc / (1 + gst_rate)
        return ex, inc - ex, inc

    def from_ex(ex):
        gst = ex * gst_rate
        return ex, gst, ex + gst

    if shape == "tbc":
        return {"mode": "tbc"}

    if shape == "indicative":
        if args.price_low is not None and args.price_high is not None:
            return {"mode": "range", "low": args.price_low, "high": args.price_high}
        if args.price_inc_gst is not None:
            ex, gst, inc = from_inc(args.price_inc_gst)
        elif args.price_ex_gst is not None:
            ex, gst, inc = from_ex(args.price_ex_gst)
        else:
            sys.exit("indicative shape needs --price-inc-gst, --price-ex-gst, "
                     "or --price-low and --price-high")
        return {"mode": "point", "ex": ex, "gst": gst, "inc": inc, "indicative": True}

    # fixed / finalised
    if args.price_inc_gst is not None:
        ex, gst, inc = from_inc(args.price_inc_gst)
    elif args.price_ex_gst is not None:
        ex, gst, inc = from_ex(args.price_ex_gst)
    else:
        sys.exit(f"{shape} shape needs --price-inc-gst or --price-ex-gst")
    return {"mode": "point", "ex": ex, "gst": gst, "inc": inc, "indicative": False}


def price_section(doc, shape, prices):
    add_heading(doc, "4. Price impact", level=1)

    if prices["mode"] == "tbc":
        add_para(
            doc,
            "Price: TO BE CONFIRMED. No amount is payable under this notice. "
            "This notice is issued now so that the variation is recorded in writing "
            "before the work proceeds, as required by section 7AAA of the Home Building "
            "Act 1989 (NSW). The price will be confirmed in a subsequent Phase 2 "
            "(Finalised) notice once supplier pricing has been received.",
            bold=True,
        )
        return

    if prices["mode"] == "range":
        add_para(
            doc,
            "The amount below is an INDICATIVE ESTIMATE ONLY and is not a fixed price. "
            "A Phase 2 (Finalised) notice with the confirmed price will be issued once all "
            "costs are known.",
            italic=True, color=GREY,
        )
        money_table(doc, [
            ("Indicative range (inc GST)",
             f'{money(prices["low"])}  –  {money(prices["high"])}', True, "FCE4D6"),
        ])
        return

    # point estimate (indicative point, fixed, or finalised)
    if prices.get("indicative"):
        add_para(
            doc,
            "The figures below are an INDICATIVE ESTIMATE ONLY and are not a fixed price. "
            "A Phase 2 (Finalised) notice with the confirmed price will follow.",
            italic=True, color=GREY,
        )
        total_fill = "FCE4D6"
        total_label = "INDICATIVE TOTAL (inc GST)"
    elif shape == "finalised":
        add_para(
            doc,
            "The following is the FINAL PRICE for this variation. All figures shown are the "
            "price payable by the Client.",
        )
        total_fill = "E2EFDA"
        total_label = "TOTAL (inc GST)"
    else:  # fixed
        add_para(
            doc,
            "The following is a FIXED PRICE for this variation, awaiting the Client's "
            "acknowledgment. All figures shown are the price payable by the Client.",
        )
        total_fill = "E2EFDA"
        total_label = "TOTAL (inc GST)"

    money_table(doc, [
        ("Sub-total (ex GST)", money(prices["ex"]), False, None),
        ("GST (10%)", money(prices["gst"]), False, None),
        (total_label, money(prices["inc"]), True, total_fill),
    ])


def contract_sum_section(doc, shape, prices, original, previous_approved):
    add_heading(doc, "5. Effect on contract sum", level=1)

    orig_str = money(original) if original is not None else "TBC"
    prev_str = money(previous_approved)

    if prices["mode"] == "tbc":
        this_str, revised_str = "TBC", "TBC"
    elif prices["mode"] == "range":
        this_str = f'{money(prices["low"])} – {money(prices["high"])} (indicative)'
        revised_str = "TBC (depends on final price)"
    else:
        this_inc = prices["inc"]
        if prices.get("indicative"):
            this_str = f'{money(this_inc)} (indicative)'
            revised_str = "TBC (depends on final price)"
        else:
            this_str = money(this_inc)
            revised_str = (money(original + previous_approved + this_inc)
                           if original is not None else "TBC")

    money_table(doc, [
        ("Original contract sum", orig_str, False, None),
        ("Plus: previously approved variations", prev_str, False, None),
        ("Plus / minus: this variation", this_str, False, None),
        ("REVISED CONTRACT SUM", revised_str, True,
         "E2EFDA" if revised_str.startswith("$") else "FCE4D6"),
    ])


# ---------------------------------------------------------------- build

def build(args, cfg, project, prices):
    b = cfg.get("builder", {})
    shape_meta = SHAPES[args.shape]

    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2.0); section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.2); section.right_margin = Cm(2.2)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)

    # Header (logo if present, else text)
    logo = next((REPO_ROOT / "templates" / f"logo.{e}"
                 for e in ("png", "jpg", "jpeg")
                 if (REPO_ROOT / "templates" / f"logo.{e}").exists()), None)
    if logo:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(logo), width=Cm(6.0))
    else:
        add_heading(doc, "RENOVATE 8", level=0)
    add_para(doc, b.get("trading_as_line", ""), bold=True, color=GREY,
             align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, f'{b.get("address","")}    |    ABN {b.get("abn","")}    |    '
                  f'NSW Contractor Licence {b.get("licence","")}',
             italic=True, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, size=9)

    # Title + banner
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    r = p.add_run("FORMAL VARIATION NOTICE"); r.bold = True
    r.font.size = Pt(16); r.font.color.rgb = NAVY

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(shape_meta["banner"]); r.bold = True
    r.font.size = Pt(12); r.font.color.rgb = shape_meta["banner_color"]

    add_para(doc,
             "Issued under the Master Builders Association of NSW Residential Building "
             "Contract and section 7AAA of the Home Building Act 1989 (NSW).",
             italic=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY)

    # Identification
    t = doc.add_table(rows=0, cols=2)
    t.autofit = False
    t.columns[0].width = Cm(5.5); t.columns[1].width = Cm(11.5)
    add_field_row(t, "Project / site address", project.get("project_name", ""))
    add_field_row(t, "Client (owner)", project.get("client", ""))
    add_field_row(t, "Head contract dated", fmt_date(project.get("contract_date", "")))
    add_field_row(t, "Variation number", args.var)
    add_field_row(t, "Notice version", f"v{args.version}")
    add_field_row(t, "Date issued", fmt_date(args.date_issued))
    add_field_row(t, "Date variation identified", fmt_date(args.date_identified))
    add_field_row(t, "Trigger / source", args.trigger)

    # 1. Nature
    add_heading(doc, "1. Nature of variation", level=1)
    add_para(doc, args.description)
    if args.reason:
        add_heading(doc, "2. Reason for variation", level=1)
        add_para(doc, args.reason)

    # 3. Scope
    add_heading(doc, "3. Scope of works", level=1)
    add_para(doc, f"Inclusions: {args.scope_inclusions or '[to be detailed]'}")
    add_para(doc, f"Exclusions: {args.scope_exclusions}")
    if args.drawing_ref:
        add_para(doc, f"Drawing / specification reference: {args.drawing_ref}")

    # 4. Price (shape-dependent)
    price_section(doc, args.shape, prices)
    if args.pricing_basis:
        add_para(doc, f"Pricing basis: {args.pricing_basis}", italic=True, size=9, color=GREY)

    # 5. Effect on contract sum
    original = (args.original_sum if args.original_sum is not None
                else project.get("original_contract_sum_inc_gst"))
    contract_sum_section(doc, args.shape, prices, original, args.previous_approved)

    # 6. Time impact
    add_heading(doc, "6. Time impact", level=1)
    if str(args.time_days).upper() in ("TBC", "TBD", ""):
        add_para(doc, "Time impact: TO BE CONFIRMED. Any effect on the construction "
                      "period will be advised in the Finalised notice.")
    else:
        verb = {"extend": "extend", "reduce": "reduce", "none": "not change"}[args.time_direction]
        if args.time_direction == "none":
            add_para(doc, "This variation will not change the construction period.")
        else:
            tail = " (indicative)" if args.shape in ("tbc", "indicative") else ""
            add_para(doc, f"This variation will {verb} the construction period by "
                          f"{args.time_days} working day(s){tail}.")
    if args.revised_pc_date:
        add_para(doc, "Revised practical completion date (subject to weather and other "
                      f"contractual extensions): {fmt_date(args.revised_pc_date)}.")

    # 7. Payment
    add_heading(doc, "7. Payment", level=1)
    add_para(doc, "This variation will be invoiced as follows (one option applies):")
    add_para(doc, "[  ] Bundled into the next progress claim (preferred, in accordance "
                  "with the head contract).")
    add_para(doc, "[  ] By a separate tax invoice issued on completion of this variation.")
    terms = cfg.get("defaults", {}).get("payment_terms", "")
    if terms:
        add_para(doc, f"Payment terms: {terms}")

    # 8. Statutory warranties
    add_heading(doc, "8. Statutory warranties", level=1)
    add_para(doc, "All work the subject of this variation is performed subject to the "
                  "statutory warranties set out in Schedule 2 of the Home Building Act 1989 "
                  "(NSW), which apply notwithstanding anything in this notice.", size=9)

    # 9. Urgency
    add_heading(doc, "9. Urgency / commencement", level=1)
    add_para(doc, "[  ] Standard — Work on this variation will not commence until this "
                  "notice has been signed by the Client.", size=10)
    add_para(doc, "[  ] Urgent / safety — Work has commenced (or will commence) before "
                  "signature because it is necessary for safety, compliance with law, or to "
                  f"prevent further loss or damage. Reason: {args.urgency_reason or '____________'}.",
             size=10)

    # 10. Acknowledgment
    add_heading(doc, "10. Client acknowledgment & acceptance", level=1)
    if args.shape == "tbc":
        add_para(doc, "By signing below, the Client acknowledges that this variation has "
                      "been raised and will be carried out, and that the price will be "
                      "confirmed by a subsequent notice before becoming payable:")
    else:
        add_para(doc, "By signing below, the Client acknowledges that:")
        for bullet in [
            "the variation has been explained to them;",
            "they consent to the variation being performed;",
            "the price and time impact set out above are accepted;",
            "they have had the opportunity to ask questions and (where they wished) seek independent advice;",
            "this acceptance forms part of the head contract and is a variation under section 7AAA of the Home Building Act 1989 (NSW).",
        ]:
            pp = doc.add_paragraph(style="List Bullet")
            pp.add_run(bullet).font.size = Pt(10)

    sigt = doc.add_table(rows=2, cols=2)
    sigt.autofit = False
    sigt.columns[0].width = Cm(8.5); sigt.columns[1].width = Cm(8.5)

    def sig_cell(cell, label, name, date):
        cell.text = ""
        r = cell.paragraphs[0].add_run(label); r.bold = True; r.font.size = Pt(10)
        p2 = cell.add_paragraph(); p2.paragraph_format.space_before = Pt(18)
        p2.add_run("Signature: ______________________________________").font.size = Pt(10)
        cell.add_paragraph().add_run(f"Print name: {name}").font.size = Pt(10)
        cell.add_paragraph().add_run(f"Date: {date}").font.size = Pt(10)

    client = project.get("client", "")
    sig_cell(sigt.rows[0].cells[0], "CLIENT (Owner)", client, "________________")
    sig_cell(sigt.rows[0].cells[1], "CLIENT (Co-owner, if applicable)", "", "________________")
    sig_cell(sigt.rows[1].cells[0], "RENOVATE 8 (Builder)",
             b.get("signatory", ""), fmt_date(args.date_issued))
    sig_cell(sigt.rows[1].cells[1], "Witness (if required)", "________________", "________________")

    # Footer
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(20)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("This is a formal variation notice. Please return a signed copy to "
                  "Renovate 8 (by email reply with signed PDF, or by hand).")
    r.italic = True; r.font.color.rgb = GREY; r.font.size = Pt(9)

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f'{b.get("trading_as_line","")}  •  ABN {b.get("abn","")}  •  '
                  f'NSW Contractor Licence {b.get("licence","")}  •  '
                  f'Filed: {args.var} / v{args.version} / {shape_meta["state"]}')
    r.italic = True; r.font.color.rgb = GREY; r.font.size = Pt(8)

    return doc


def to_pdf(docx_path: Path) -> Path | None:
    """Best-effort PDF export via LibreOffice. Returns the pdf path or None.

    LibreOffice exits 0 even when it fails to load a source file, so we judge
    success purely by whether the .pdf actually appears. Some sandboxes ship a
    non-functional LibreOffice — in that case this returns None and the caller
    falls back to delivering the .docx (Word can Save-As-PDF in one click).
    """
    pdf = docx_path.with_suffix(".pdf")
    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf",
             "--outdir", str(docx_path.parent), str(docx_path)],
            capture_output=True, timeout=120,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return None
    return pdf if pdf.exists() else None


def main() -> int:
    cfg = load_config()
    gst_rate = cfg.get("defaults", {}).get("gst_rate", 0.10)

    ap = argparse.ArgumentParser(description="Generate a filled Variation Notice.")
    ap.add_argument("--project-code", required=True)
    ap.add_argument("--var", required=True, help='e.g. "V5"')
    ap.add_argument("--version", default="1", help="Notice version (default 1)")
    ap.add_argument("--shape", required=True, choices=list(SHAPES.keys()))
    ap.add_argument("--description", required=True)
    ap.add_argument("--reason", default="")
    ap.add_argument("--scope-inclusions", default="")
    ap.add_argument("--scope-exclusions", default="Nil")
    ap.add_argument("--drawing-ref", default="")
    ap.add_argument("--trigger", default="")
    ap.add_argument("--date-identified", default="")
    ap.add_argument("--date-issued", default=dt.date.today().isoformat())
    ap.add_argument("--price-inc-gst", type=float, default=None)
    ap.add_argument("--price-ex-gst", type=float, default=None)
    ap.add_argument("--price-low", type=float, default=None)
    ap.add_argument("--price-high", type=float, default=None)
    ap.add_argument("--pricing-basis", default="")
    ap.add_argument("--time-days", default="TBC")
    ap.add_argument("--time-direction", default="extend", choices=["extend", "reduce", "none"])
    ap.add_argument("--revised-pc-date", default="")
    ap.add_argument("--original-sum", type=float, default=None,
                    help="Override the project's original contract sum (inc GST)")
    ap.add_argument("--previous-approved", type=float, default=0.0,
                    help="Sum of previously approved variations (inc GST). Read from the "
                         "register and pass in; defaults to 0.")
    ap.add_argument("--urgency-reason", default="")
    ap.add_argument("--out", default=None)
    ap.add_argument("--pdf", action="store_true", help="Also export a PDF via libreoffice")
    args = ap.parse_args()

    projects = cfg.get("projects", {})
    project = projects.get(args.project_code)
    if project is None:
        sys.exit(f"Unknown project code '{args.project_code}'. "
                 f"Known: {', '.join(k for k in projects if not k.startswith('_'))}")

    prices = compute_prices(args, gst_rate)
    doc = build(args, cfg, project, prices)

    state = SHAPES[args.shape]["state"]
    stamp = dt.date.fromisoformat(args.date_issued).strftime("%y%m%d")
    if args.out:
        out = Path(args.out)
    else:
        WORKING.mkdir(exist_ok=True)
        out = WORKING / f'{args.project_code}_{args.var}_v{args.version}_{state}_{stamp}.docx'
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print(f"Wrote {out}")
    print(f"  shape={args.shape}  banner=\"{SHAPES[args.shape]['banner']}\"")

    if args.pdf:
        pdf = to_pdf(out)
        if pdf:
            print(f"Wrote {pdf}")
        else:
            print("PDF not produced (LibreOffice unavailable/non-functional here). "
                  "Deliver the .docx; export to PDF in Word (File > Save As > PDF).")

    print("\nNext: review, then deliver the .docx (and .pdf) to Paul via SendUserFile "
          f"for upload to the variation's 03_Notice/ folder as "
          f"{args.var}_v{args.version}_{state}_{stamp}.pdf")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
