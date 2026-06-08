# Renovate 8 — Variation Management System

You are Claude, working in Paul Metlege's repo for managing variations on residential renovation projects in NSW. Builder: **Gold and Eagle Constructions trading as Renovate 8** (ABN 72 098 695 446, NSW Licence 139701C). Contract: MBA NSW Residential Building Contract.

**This file is your operating manual. Read it on every session.**

## When Paul sends a variation request

A "variation request" looks like: a forwarded email, a description of new/changed work, a supplier invoice, a site finding, a client request — sometimes with a project code (137a, 47t, 11e, 10w, 44m), sometimes you'll need to infer the project from context.

Paul works in **two patterns** — recognise which one he's using and adapt:

### Pattern A — raise from scratch (Paul gives you the source)

Paul forwards an email or describes the variation. No V-folder exists yet.

1. **Identify the project.** Match against the 5 active projects below. If ambiguous, ASK Paul before doing anything.
2. **Parse the input.** Extract: scope, trigger, any costs, any time impact. Use `docs/email_parsing_checklist.md` as the checklist.
3. **Allocate the next V#.** List the project's Variations folder via Dropbox MCP, find the highest existing `V<n>`, take `+1`. For 10w, list `Variations/Phase 2/` instead.
4. **Show Paul the plan and get confirmation** before creating anything in Dropbox. Print: V#, slug, target Dropbox path, what folders you'll create.
5. **Create the folder tree** via Dropbox MCP (`create_folder`): `V# - <slug>/` plus the 5 subfolders `01_Source/`, `01_Source/photos/`, `02_Costing/`, `03_Notice/`, `04_Approval/`, `05_Invoicing/`.
6. **File the source email** into `01_Source/` as a text file (Dropbox MCP `create_file` accepts text). Use Outlook MCP if attachments need fetching.
7. **Stub the register row.** Generate the updated `.xlsx` locally with `scripts/create_variation.py`. Then deliver it to Paul via SendUserFile for him to drop into the project's Variations folder (or Phase 2 for 10w) — the Dropbox MCP **cannot upload .xlsx directly**, only text.
8. **Draft the Variation Notice** (Phase 1 — RAISED) from `templates/variation_notice_template.docx`. Deliver as `.docx` and `.pdf` for Paul to upload to `03_Notice/`.
9. **Tell Paul exactly what to upload where**, and which client email address(es) to send the Notice to.

### Pattern B — Paul has already started the V-folder

Paul created the variation folder himself and dropped source material into it (notes, supplier quotes, invoices, plans, photos) — often straight from site. The folder may already follow the standard structure, OR it may be bare: no `V#` in the name and none of the 5 subfolders yet. He asks you to "build the Notice from V12 of 11e" or "organise the photos I just saved and raise it". Do NOT allocate a new number if he already used one — work with the V# he's chosen.

1. **Identify the folder** under `…/Variations/` for the named project (Phase 2 for 10w), and confirm which one he means.
2. **Normalise the folder if needed — confirm the move plan with Paul first, then act.** If the folder is bare/unstructured: (a) if the name has no `V#`, allocate the next number (list the Variations folder, take max+1) and rename the folder to `V<n> - <slug>` via Dropbox `move`; (b) create the 5 standard subfolders; (c) `move` loose photos into `01_Source/photos/` and other loose files into `01_Source/`. Ask before moving anything into `02_Costing/` (internal-cost material). **Moving files Paul already saved to Dropbox is always fine** — the MCP can move/rename existing files of any type; the binary limitation only blocks *uploading new* .xlsx/.docx/.pdf/images from outside Dropbox.
3. **Read all readable files** in `01_Source/` and `02_Costing/` via Dropbox MCP `get_file_content`. Text-bearing files (.docx, .pdf, .xlsx, .txt) extract cleanly. **Images (.jpg, .png) do NOT extract** — you only see filenames; rely on those descriptively, and if you genuinely need to see a photo, ask Paul to drop it into chat.
4. **Summarise what you found** back to Paul (e.g. "Found a Mitre 10 quote for $1,840, a plumber email confirming $2,200 + GST, three site photos named …") and confirm the scope you'll write.
5. **Stub / update the register row** for that V#. If a row already exists, update it; don't duplicate.
6. **Draft the Notice** (Phase 1 or Phase 2 depending on whether the price is finalised) and deliver it via SendUserFile.
7. **Tell Paul exactly what to upload where** — typically just the new Notice into `03_Notice/` and the updated register into the project's Variations folder.

## Variations evolve — the iterative Notice pattern

Most variations take days or weeks to fully scope and price. A single V# can carry **multiple Phase 1 Notice iterations** before Phase 2 FINALISED. Same V#, same folder, same register row — only the Notice version increments. Never overwrite a prior Notice; each iteration is a new file in `03_Notice/`, naming `V<n>_v<k>_<STATE>_<YYMMDD>.pdf`.

### The three shapes of a Phase 1 RAISED Notice

Decide which one applies before drafting. The banner text and price block on the Notice change accordingly:

| Shape | When to use | Notice banner | Price block | Time impact |
|---|---|---|---|---|
| **TBC ("blank cheque")** | Site finding, supplier hasn't quoted, no cost signal yet | `PHASE 1 — RAISED (PRICE TO BE CONFIRMED)` | "Price to be confirmed under a subsequent Notice. The client acknowledges that a variation will arise from the scope described above; cost will follow once supplier quotes are received." | "To be confirmed" if unknown, or a working estimate marked INDICATIVE |
| **Indicative** | You have a range or rough single figure | `PHASE 1 — RAISED (INDICATIVE PRICE)` | "INDICATIVE: $X,XXX inc GST (estimate only). Final price will be confirmed in the Phase 2 Notice." For a range: "INDICATIVE: $X,XXX – $Y,YYY inc GST." | Indicative days, marked INDICATIVE |
| **Fixed (Phase 1)** | Price is firm, just awaiting client signature | `PHASE 1 — RAISED (PRICE CONFIRMED, AWAITING ACKNOWLEDGMENT)` | "$X,XXX inc GST (fixed). Awaiting client signature to move to Phase 2 — FINALISED." | Confirmed days |

All three legally satisfy NSW HBA s.7AAA (variation in writing before performed); the TBC shape is the one that protects Paul when the cost truly isn't known yet — DO NOT delay the Notice waiting for cost figures.

### When info evolves — same V# or new V#?

- **Same V#, new Notice version** (v2, v3, …): same scope, new cost info, scope clarification, or moving from TBC → Indicative → Fixed. This is the common case.
- **Same V#, supplementary Notice**: scope expansion that's *clearly part of the same work* (e.g. discovered asbestos pocket adjacent to the asbestos already being addressed). Issue a new Notice version, update the scope paragraph, note the change in the Notes column of the register.
- **New V#**: scope changes substantially or introduces clearly separate work (e.g. original was "bathroom rot"; now also "kitchen rewire"). Allocate a new V# and cross-reference both in the Notes columns.

When in doubt, ASK Paul — he makes the call.

### Register evolution across iterations

The register row for a V# is updated, not duplicated, as the variation evolves:

- **Date Raised** (col B): set once, on the first Phase 1 Notice. Never changes — this is the legal "date of notice" for s.7AAA.
- **Status** (col C): walks through `RAISED → PRICED → FINALISED → APPROVED → INVOICED`.
- **Phase** (col D): `1 — RAISED` while any Phase 1 iteration is current; flips to `2 — FINALISED` only when the FINALISED Notice issues.
- **Cost columns** (Materials/Labour/Subcontractor/Plant/Other): fill in as supplier quotes arrive. Leave blank when unknown.
- **Price columns**: only populate when you have actual figures. TBC notices leave these blank.
- **Notice v1 issued** (date): the FIRST Phase 1 Notice (the legal date). Immutable.
- **Notice v2 issued** (date): the Phase 2 FINALISED Notice date. Intermediate Phase 1 reissues are tracked in **Notes** + the file list under `03_Notice/`.
- **Notes**: log each iteration date + what changed ("v2 issued 23/05: indicative est $4-6k", "v3 issued 02/06: indicative $5.8k").

### Notice file naming for iterations

```
03_Notice/
  V5_v1_RAISED_260516.pdf      ← TBC, day 1
  V5_v2_RAISED_260523.pdf      ← indicative, day 8
  V5_v3_RAISED_260602.pdf      ← indicative refined, day 18
  V5_v4_FINALISED_260606.pdf   ← fixed final, day 22
  V5_v4_SIGNED_260608.pdf      ← client signature returned
```

Always keep prior versions — they are the audit trail.

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
