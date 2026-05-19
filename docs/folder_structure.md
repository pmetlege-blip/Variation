# Dropbox Folder Structure

The system works **inside** the existing project structure Renovate 8 already
uses — no new top-level folders are created.

## Top-level (already exists)

```
/Projects/
└── [Suburb] - [Address]/                       e.g. "Annandale - 137 Annandale Street"
    ├── Construction/
    ├── Contract/
    ├── Client/
    ├── Council/
    ├── Consultant/
    ├── HOWI/
    ├── Handover/
    ├── Progress Claim/
    ├── Programme/
    ├── Site/
    ├── Tender Information/
    └── Variations/                              ← all variation work lives here
```

(Folder names vary slightly across older projects; the system reads the
project folder by name rather than assuming a fixed list.)

## Inside `Variations/`

```
Variations/
├── Variation Register.xlsx                     ← per-project master register
├── Archive/                                    ← (existing) cancelled / obsolete
├── V1/                                         ← existing
├── V2 - Asbestos under Concrete/               ← existing
├── V3 - Absestos Removal/                      ← existing
├── V4 - Bamboo Works and water detailing/      ← existing
├── V5 - [description]/                         ← created by the system
├── V6 - [description]/                         ← ...
└── ...
```

## Per-variation folder layout

Inside each `V# - [description]/`:

```
V# - [description]/
├── 01_Source/                  what came in
│   ├── original_email.eml      forwarded email from Paul or client
│   ├── invoice_<supplier>.pdf  supplier invoices
│   ├── quote_<supplier>.pdf    supplier quotes
│   ├── photos/                 site photos
│   └── plans_markup.pdf        marked-up drawings if applicable
├── 02_Costing/                 CONFIDENTIAL
│   └── cost_breakdown.xlsx     internal cost-to-me workings
├── 03_Notice/                  what the client sees
│   ├── V#_v1_RAISED.docx       Phase 1 (price TBC or indicative)
│   ├── V#_v1_RAISED.pdf
│   ├── V#_v2_FINALISED.docx    Phase 2 (confirmed price)
│   ├── V#_v2_FINALISED.pdf
│   └── V#_SIGNED.pdf           fully signed copy
├── 04_Approval/                proof of acceptance under HBA s.7AAA
│   └── client_approval_email.eml
└── 05_Invoicing/               money trail
    ├── claim_reference.txt     which progress claim / invoice it went on
    └── invoice_to_client.pdf
```

## Naming rules

- **Project folder** — already named by Paul. The system reads it as-is.
  Example: `Annandale - 137 Annandale Street`.
- **V number** — `V` + integer, **no zero-padding**, matching the existing
  Renovate 8 convention (`V1`, `V2`, …, `V10`, `V11`, ...).
- **Slug** — short plain-English description, hyphens/spaces OK. Matches
  the existing pattern (e.g. `V2 - Asbestos under Concrete`).
- **Numbering is per-project.** Reset to V1 for each new project.
- **No gaps.** Even cancelled variations keep their number; the
  `Archive/` subfolder can hold cancelled or duplicate folders if needed.

## Master register

Lives at:

```
/Projects/[Suburb] - [Address]/Variations/Variation Register.xlsx
```

One register per project. Sits in the same folder as the V# folders so
opening either gets you to the other in one click.

Existing Excel files in the Variations folder (e.g. `137a.Payment
Milestone Schedule.Variation.YYMMDD.xlsx`) are left untouched.

## Where the per-variation source documents go

When Paul forwards an email with an invoice:

1. The forwarded email → `01_Source/original_email.eml`
2. Attached invoices → `01_Source/invoice_<supplier>.pdf`
3. Attached photos → `01_Source/photos/`
4. Markup drawings → `01_Source/`
5. Internal cost workings → `02_Costing/cost_breakdown.xlsx` (stays out of
   the client's hands; Dropbox permissions can be tightened on this
   subfolder if the team grows)
6. The formal notice → `03_Notice/` (this is what gets emailed to the client)
7. The signed-and-returned notice + approval email → `04_Approval/`
8. The progress-claim line reference or standalone invoice → `05_Invoicing/`
