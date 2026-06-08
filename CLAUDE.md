# Renovate 8 — Variation Management System

You are Claude, working in Paul Metlege's repo for managing variations on residential renovation projects in NSW. Builder: **Gold and Eagle Constructions trading as Renovate 8** (ABN 72 098 695 446, NSW Licence 139701C). Contract: MBA NSW Residential Building Contract.

**This file is your operating manual. Read it on every session.**

## When Paul sends a variation request

A "variation request" looks like: a forwarded email, a description of new/changed work, a supplier invoice, a site finding, a client request — sometimes with a project code (137a, 47t, 11e, 10w, 44m), sometimes you'll need to infer the project from context.

When that happens, follow this workflow:

1. **Identify the project.** Match against the 5 active projects below. If ambiguous, ASK Paul before doing anything.
2. **Parse the email.** Extract: scope, trigger, any costs, any time impact. Use `docs/email_parsing_checklist.md` as the checklist.
3. **Allocate the next V#.** List the project's Variations folder via Dropbox MCP, find the highest existing `V<n>`, take `+1`. For 10w, list `Variations/Phase 2/` instead.
4. **Show Paul the plan and get confirmation** before creating anything in Dropbox. Print: V#, slug, target Dropbox path, what folders you'll create.
5. **Create the folder tree** via Dropbox MCP (`create_folder`): `V# - <slug>/` plus the 5 subfolders `01_Source/`, `01_Source/photos/`, `02_Costing/`, `03_Notice/`, `04_Approval/`, `05_Invoicing/`.
6. **File the source email** into `01_Source/` as a text file (Dropbox MCP `create_file` accepts text). Use Outlook MCP if attachments need fetching.
7. **Stub the register row.** Generate the updated `.xlsx` locally with `scripts/create_variation.py`. Then deliver it to Paul via SendUserFile for him to drop into the project's Variations folder (or Phase 2 for 10w) — the Dropbox MCP **cannot upload .xlsx directly**, only text.
8. **Draft the Variation Notice** (Phase 1 — RAISED) from `templates/variation_notice_template.docx`. Deliver as `.docx` and `.pdf` for Paul to upload to `03_Notice/`.
9. **Tell Paul exactly what to upload where**, and which client email address(es) to send the Notice to.

When the price is later finalised, repeat steps 7–9 but for Phase 2 — FINALISED.

## Active projects (source of truth: `scripts/config.json`)

| Code | Dropbox folder | Client | Variations live in | Contract sum (inc GST) |
|---|---|---|---|---|
| 137a | `Annandale - 137 Annandale Street` | Jon and Stacey Day | `Variations/` | $2,997,386.00 |
| 47t | `Haberfield - 47 Tillock Street` | Joanne King | `Variations/` | $435,000.00 |
| 11e | `Lilyfield - 11 Eric Street` | Nicholas Lewis and Sarah Johnson | `Variations/` | $722,790.00 |
| 10w | `Taren Point - 10 Woodlands Road` | Chris and Joanna Gialouris | **`Variations/Phase 2/`** ⚠️ | TBC |
| 44m | `Sylvania Waters - 44 Macintyre Crescent` | David and Vivian Salim | `Variations/` | $1,290,601.40 |

**10w special rule:** all variation work goes in `Variations/Phase 2/`. The `Phase 1/` subfolder is historical — **never create in or modify Phase 1**. `create_variation.py --project-code 10w` handles this automatically.

## Tools you have

- **Dropbox MCP** — read, list, search, create folders, move/copy/delete. Cannot upload binary files (.xlsx, .docx, .pdf): deliver those to Paul via `SendUserFile` for him to drop in.
- **Outlook MCP** — read/search Paul's emails and attachments if needed.
- **GitHub MCP** — limited to the `pmetlege-blip/variation` repo only.
- **Python scripts** in `scripts/`:
  - `create_variation.py` — allocates V#, plans the folder tree, stubs a register row. Takes `--project-code` to honor per-project rules (e.g. 10w → Phase 2).
  - `setup_project_register.py` — one-time, generates a project's master register.
  - `build_register_template.py` and `build_notice_template.py` — regenerate the .xlsx and .docx templates from `config.json`.

## Rules — non-negotiable

- **Never write the cost-to-me figures onto the client-facing Notice.** Internal costs live in `02_Costing/` only. The Notice shows price-to-client only.
- **Never modify 10w/Variations/Phase 1.**
- **Always confirm with Paul before any Dropbox mutation** (create_folder, move, delete). Read-only ops (list, get_file_content) don't need confirmation.
- **Do not create or modify the variation register** without showing Paul the new contents first.
- **Two figures Paul supplies, you do not search for:** contract date and original contract sum. Everything else is in the Write Up (`<project>/Tender/Write Up/` or `<project>/Tender Information/Write Up/`) or the Progress Claim invoices.

## Docs to read for detail

- `README.md` — overview
- `docs/workflow.md` — two-phase workflow (RAISED → FINALISED → APPROVED → INVOICED), state machine, what each phase contains
- `docs/folder_structure.md` — Dropbox layout per project and per variation
- `docs/numbering_scheme.md` — V# scheme, gaps, versioning
- `docs/email_parsing_checklist.md` — fields to extract from each variation email
- `docs/legal_compliance.md` — NSW Home Building Act s.7AAA requirements
- `docs/project_setup.md` — one-time per-project setup

## Branch

All work is committed to `claude/variation-management-system-fWlUg`. Don't push to other branches without Paul's explicit say-so.
