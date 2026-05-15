# Dropbox Folder Structure

All variation documents live in Dropbox. Structure is **strict** so that the
register, the notice, and the source documents can always be cross-referenced.

## Top-level

```
/Renovate 8/
└── Projects/
    └── [PROJECT_CODE] - [Client surname or site address]/
        ├── 00_Contract/
        ├── 01_Plans_and_Specs/
        ├── 02_Progress_Claims/
        ├── 03_Variations/
        │   ├── _Register/
        │   │   └── [PROJECT_CODE]_Variation_Register.xlsx     ← master register
        │   ├── [PROJECT_CODE]-VAR-001 - [short slug]/
        │   ├── [PROJECT_CODE]-VAR-002 - [short slug]/
        │   └── ...
        ├── 04_Photos/
        ├── 05_Correspondence/
        └── 99_Archive/
```

## Per-variation folder

Inside each `[PROJECT_CODE]-VAR-###` folder:

```
[PROJECT_CODE]-VAR-### - [short slug]/
├── 01_Source/
│   ├── original_email.eml          forwarded email from Paul (or client)
│   ├── invoice_<supplier>.pdf      any supplier invoices
│   ├── quote_<supplier>.pdf        any supplier quotes
│   ├── photos/                     site photos (before / during / after)
│   └── plans_markup.pdf            marked-up drawings if applicable
├── 02_Costing/
│   ├── cost_breakdown.xlsx         internal cost-to-me workings
│   └── markup_calc.xlsx            (optional) markup decision rationale
├── 03_Notice/
│   ├── VAR-###_v1_RAISED.docx      Phase 1 notice (price TBC or indicative)
│   ├── VAR-###_v1_RAISED.pdf       PDF of the above (what is sent)
│   ├── VAR-###_v2_FINALISED.docx   Phase 2 notice (confirmed price)
│   ├── VAR-###_v2_FINALISED.pdf
│   └── VAR-###_SIGNED.pdf          fully signed copy returned by client
├── 04_Approval/
│   ├── client_approval_email.eml   written approval (HBA s.7AAA evidence)
│   └── any_back_and_forth.eml      negotiation correspondence
└── 05_Invoicing/
    ├── claim_reference.txt         which progress claim or invoice it went on
    └── invoice_to_client.pdf       if invoiced separately
```

## Naming rules

- **PROJECT_CODE** — short uppercase identifier, no spaces. Either:
  - the street number + first 5 letters of street/surname (e.g. `42SMITH`,
    `17BRIDG`), or
  - an explicit job code if the contract uses one.
- **VAR number** — zero-padded to 3 digits. Reset to `001` for each new project.
- **Short slug** — 3–6 words, kebab-case, plain English. Examples:
  - `extra-power-point-kitchen`
  - `engineer-redesign-footings`
  - `client-changed-tile-spec`
- Never rename or delete a folder once created. If a variation is cancelled,
  add the suffix ` - CANCELLED` and update the register status.

## Why this structure

- **`01_Source`** isolates "what came in" so the audit trail is preserved.
- **`02_Costing`** is **confidential** — internal cost-to-me workings live
  here and never go to the client. (Dropbox folder permissions on this
  sub-folder can be restricted further if the project team grows.)
- **`03_Notice`** holds the documents the **client actually sees**.
- **`04_Approval`** isolates the s.7AAA written-evidence that the client
  accepted the variation. This is the single most legally important folder.
- **`05_Invoicing`** ties the variation to the money trail.

## Master register location

The per-project master register lives at:

```
/Renovate 8/Projects/[PROJECT_CODE] - .../03_Variations/_Register/[PROJECT_CODE]_Variation_Register.xlsx
```

There is **no global cross-project register** by design — keep books per-project
so that a single client's documents can be exported cleanly (and to keep
cost-to-me data scoped to one site at a time).
