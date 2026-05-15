# Renovate 8 — Variation Management System

A complete workflow for managing variations on residential renovation projects
in NSW, compliant with the **Home Building Act 1989 (NSW)** and the
**Master Builders Association of NSW Residential Building Contract**.

## What this system does

1. **Receives** variation requests via email (from Paul Metlege or directly from clients/site).
2. **Parses** the email to extract project, scope, trigger, and known costs.
3. **Numbers** the variation sequentially per project (e.g. `42SMITH-VAR-007`).
4. **Creates** a dedicated Dropbox folder for the variation under the project.
5. **Files** invoices, quotes, photos, emails, and the formal notice in the folder.
6. **Adds** the variation to the project's master Excel register.
7. **Issues** a formal NSW-compliant Variation Notice to the client:
   - **Phase 1 — Raised:** description + scope + indicative price (or TBC) + time impact.
   - **Phase 2 — Finalised:** confirmed price + signed acknowledgment + new contract sum.
8. **Keeps** Cost-to-me (internal) separate from Price-to-client (on the notice).

## Repository layout

```
Variation/
├── README.md                            this file
├── docs/
│   ├── folder_structure.md              Dropbox layout per project / variation
│   ├── workflow.md                      Two-phase workflow + diagram
│   ├── email_parsing_checklist.md       What to extract from each email
│   ├── numbering_scheme.md              How variations are numbered
│   └── legal_compliance.md              NSW HBA + MBA NSW contract notes
├── templates/
│   ├── master_register_template.xlsx    Per-project Excel register
│   └── variation_notice_template.docx   Word notice (Phase 1 + Phase 2)
├── scripts/
│   ├── build_register_template.py       Regenerates the .xlsx template
│   ├── build_notice_template.py         Regenerates the .docx template
│   ├── create_variation.py              Helper: provision a new variation
│   └── config.example.json              Business details (license, ABN, etc.)
└── examples/
    └── sample_email_intake.md           Worked example of email -> variation
```

## Quick start (operator)

When Paul forwards an email containing an invoice or variation request:

1. Open `docs/email_parsing_checklist.md` and tick through each field.
2. If price is missing or unclear, **reply asking Paul** for the missing info
   (see the checklist for the standard question set).
3. Run `scripts/create_variation.py --project 42SMITH --description "..."`
   to allocate the next number and create the Dropbox folder.
4. Save all source documents (invoices, photos, quotes, the original email)
   into the new folder.
5. Add a row to the project's master register.
6. Generate the Phase 1 Variation Notice from `variation_notice_template.docx`.
7. Email the notice to the client for acknowledgment.
8. Once costs are finalised, regenerate the notice as **Phase 2 — Finalised**,
   update the register, and re-issue.

## Start here

- New to the system? Read `docs/workflow.md` first.
- Setting up a new project? Copy `templates/master_register_template.xlsx`
  into the project's Dropbox folder and rename it.
- Need to check compliance? See `docs/legal_compliance.md`.
