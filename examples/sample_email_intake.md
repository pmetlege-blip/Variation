# Worked example — email intake to issued notice

This walks through a realistic scenario end-to-end so a new operator can
see how all the pieces fit together.

## Step 1 — The inbound email

> **From:** Paul Metlege <paul@renovate8.example>
> **To:** [you]
> **Subject:** Fwd: extra power points 42 Smith kitchen
> **Date:** Mon, 11 May 2026 14:32
>
> Hi,
>
> Forwarding the sparky's invoice. Smiths asked for 4 extra GPOs in the
> kitchen during the rough-in walkthrough on Friday. Spark did it the same
> day so we didn't hold up the plasterer.
>
> Invoice attached — $480 inc GST from Volt Electrical. My labour to
> coordinate was probably an hour. No materials beyond what's on the
> invoice.
>
> Apply standard markup.
>
> — Paul
>
> *(Attached: volt_invoice_8842.pdf — $436.36 ex GST + $43.64 GST = $480.00)*

## Step 2 — Parse the email (using `docs/email_parsing_checklist.md`)

| Field | Value |
|---|---|
| Project | `42SMITH` (from "42 Smith kitchen") |
| Trigger | Client request (asked during walkthrough) |
| Description | Extra GPOs in kitchen |
| Detailed scope | Install 4 additional general power outlets (GPOs) in kitchen during rough-in. By Volt Electrical. Same circuit as existing kitchen GPOs. |
| Materials $ | (on invoice) |
| Labour $ | Paul's coordination, ~1 hr — use rate from project setup |
| Subcontractor $ | $436.36 ex GST (from invoice) |
| Plant/hire | $0 |
| Other | $0 |
| Markup | Paul said "standard" — but the system has no default. **Ask Paul.** |
| Time impact | 0 days (done same day) |
| Urgent? | No, but already performed — flag as "papered up after the fact" |

## Step 3 — Reply to Paul (markup is missing)

> Hi Paul,
>
> Picking up the extra GPOs at 42 Smith. Volt invoice is clear ($436.36
> ex GST). Quick question:
>
> 1. **Markup** — what % for this one? (System has no default — set per-variation.)
> 2. **Your labour rate** — I'll log 1 hr of your time at the rate on
>    the 42SMITH project setup, please confirm that's still current.
>
> Once you confirm I'll raise **42SMITH-VAR-007** (Phase 1) and get the
> notice to the Smiths. Time impact is zero (Volt did it the same day) so
> Phase 2 will follow same-day.
>
> Cheers,

(Paul replies: *"20% markup. Hourly rate $95 ex GST. Cheers."*)

## Step 4 — Allocate the variation

```bash
python3 scripts/create_variation.py \
    --project 42SMITH \
    --description "extra GPOs in kitchen" \
    --trigger "Client request"
```

Output:
```
Variation ID:    42SMITH-VAR-007
Slug:            extra-gpos-in-kitchen
Trigger:         Client request
Register:        working/42SMITH_Variation_Register.xlsx
Dropbox plan:
   create folder: /Renovate 8/Projects/42SMITH - .../03_Variations/42SMITH-VAR-007 - extra-gpos-in-kitchen
   create folder: ... /01_Source
   create folder: ... /01_Source/photos
   create folder: ... /02_Costing
   create folder: ... /03_Notice
   create folder: ... /04_Approval
   create folder: ... /05_Invoicing

Stubbed register row 4 in working/42SMITH_Variation_Register.xlsx (status=RAISED, phase=1).
```

(Claude then creates those folders in Dropbox via the Dropbox MCP tools.)

## Step 5 — Save source documents

- Save the forwarded email as `original_email.eml` in `01_Source/`.
- Save `volt_invoice_8842.pdf` into `01_Source/`.
- Open `02_Costing/cost_breakdown.xlsx` and record:
  - Subcontractor: $436.36 ex GST (Volt invoice)
  - Labour (Paul): 1 hr × $95 = $95.00 ex GST
  - Cost to me: $531.36 ex GST
  - Markup: 20% → margin $106.27
  - Price ex GST: $637.63 → GST $63.76 → inc GST $701.40

## Step 6 — Update the register

In `42SMITH_Variation_Register.xlsx` row 4:
- **J** Subcontractor $: 436.36
- **I** Labour $: 95.00
- **N** Markup %: 20%

The formulas populate:
- **M** Cost-to-me total: $531.36
- **O** Price ex GST: $637.63
- **P** GST: $63.76
- **Q** Price inc GST: $701.40
- **S** Margin: $106.27

Set **R** Time impact: 0.

## Step 7 — Issue the Phase 1 notice

In this case, costs are already known (sub invoice in hand, markup
confirmed, time impact zero) — so the Phase 1 notice is issued with a
**confirmed price**, not TBC.

Open `templates/variation_notice_template.docx`, save as
`42SMITH-VAR-007/03_Notice/VAR-007_v1_RAISED.docx`, and fill:

| Placeholder | Value |
|---|---|
| `{{STATUS_BANNER}}` | `PHASE 1 — RAISED (PRICE CONFIRMED — AWAITING ACKNOWLEDGMENT)` |
| `{{PROJECT_ADDRESS}}` | 42 Smith Street, Bondi NSW 2026 |
| `{{CLIENT_NAME}}` | John & Jane Smith |
| `{{CONTRACT_DATE}}` | 2026-01-15 |
| `{{VAR_NUMBER}}` | 42SMITH-VAR-007 |
| `{{NOTICE_VERSION}}` | 1 |
| `{{DATE_ISSUED}}` | 2026-05-11 |
| `{{DATE_TRIGGER}}` | 2026-05-08 |
| `{{TRIGGER}}` | Client request (raised during rough-in walkthrough) |
| `{{DESCRIPTION_SHORT}}` | Install 4 additional general power outlets (GPOs) in the kitchen. |
| `{{REASON}}` | The Client identified additional appliance positions during the rough-in walkthrough on 8 May 2026 and requested four additional GPOs above the bench. |
| `{{SCOPE_INCLUSIONS}}` | Supply and install 4 × double GPOs by Volt Electrical, on the existing kitchen circuit; certificate of compliance provided. |
| `{{SCOPE_EXCLUSIONS}}` | No changes to switchboard. No additional circuits. |
| `{{PRICE_DIRECTION}}` | increase |
| `{{PRICE_EX_GST}}` | $637.63 |
| `{{PRICE_GST}}` | $63.76 |
| `{{PRICE_INC_GST}}` | $701.40 |
| `{{PRICING_BASIS}}` | Subcontractor invoice (Volt Electrical #8842) plus builder coordination labour, with 20% margin. |
| `{{TIME_DIRECTION}}` | not change |
| `{{TIME_DAYS}}` | 0 |
| `{{REVISED_PC_DATE}}` | Unchanged from current programmed date |
| `{{CS_ORIGINAL}}` | $850,000.00 |
| `{{CS_PREVIOUS_VARS}}` | $12,430.00 *(sum of approved VAR-001 through VAR-006)* |
| `{{CS_THIS_VAR}}` | $701.40 |
| `{{CS_NEW_TOTAL}}` | $863,131.40 |
| `{{PAYMENT_TERMS}}` | Payable on issue of next progress claim. |
| `{{URGENCY_REASON}}` | N/A — standard. *(The "[X] Standard" box is ticked, even though work was performed before sign-off, because no urgency carve-out applies. Flag this in the cover email.)* |
| `{{CLIENT_NAME_2}}` | *(if Jane signs separately, list her)* |
| `{{BUILDER_SIGNATORY}}` | Paul Metlege |
| `{{BUILDER_ABN}}` | (from config) |
| `{{BUILDER_LICENCE}}` | (from config) |

Export to PDF, save in same folder.

## Step 8 — Email to client

> **To:** john.smith@example.com; jane.smith@example.com
> **Subject:** Variation 42SMITH-VAR-007 — extra power points (for signature)
>
> Hi John and Jane,
>
> Following the rough-in walkthrough on Friday where you asked for the
> extra GPOs in the kitchen, please find attached the formal variation
> notice — 42SMITH-VAR-007. Volt did the work the same day so the
> plasterer wasn't held up.
>
> Total cost: **$701.40 inc GST**. No change to the program.
>
> Could you please sign and return at your convenience — reply email
> with a signed PDF is fine.
>
> Cheers,
> Paul / Renovate 8

Save the sent email to `04_Approval/issued_to_client.eml`.

## Step 9 — Client returns signed PDF

- Save into `04_Approval/VAR-007_SIGNED.pdf`.
- Update register row 4:
  - **V** Client approval date: 2026-05-12
  - **W** Approval method: Signed PDF
  - **X** Approved by: John Smith
  - **C** Status: APPROVED

## Step 10 — Invoice

When next progress claim is issued, add VAR-007 line item ($701.40 inc
GST). Then update register:
  - **Y** Invoice/claim #: PC-005
  - **Z** Invoice method: Next progress claim
  - **AA** Date invoiced: 2026-05-20
  - **C** Status: INVOICED

The Summary sheet now shows:
- Variations APPROVED+INVOICED total: $13,131.40 inc GST
- Revised contract sum: $863,131.40
- (Internal, never shared) cost-to-me, margin, average markup.

Done. Folder contains the complete audit trail; register tells the
project story; client has a signed notice.
