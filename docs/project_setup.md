# Project Setup (one-time, per project)

Run this once per project to mint the master register and drop it into the
project's existing `Variations/` folder. After this, the project is "live"
and variations can be raised against it.

## Inputs

Two figures are **supplied by Paul** — do **not** go searching for them, they
are faster and more authoritative coming straight from him:

| Input | Source |
|---|---|
| **Contract date** | Provided by Paul |
| **Original Contract Sum (inc GST)** | Provided by Paul |

Everything else is derived automatically (see below).

## Client details — default lookup: the Write Up

The client name(s) and project address are read from the project's **Write Up**
document. This is the canonical source.

**Default path:**

```
/Projects/<Suburb> - <Address>/Tender Information/Write Up/
    R8.<Suburb><Address>.WriteUp.<YYMMDD>.docx     ← use the latest dated version
```

The Write Up is a Word `.docx`, so its text is readable directly via the
Dropbox MCP (`get_file_content`) — unlike scanned contract PDFs, which are
images with no extractable text.

When several dated versions exist, **use the most recent** (highest `YYMMDD`
in the filename), and prefer a copy `Amended by PM` if present.

**Fallback source:** if the Write Up is missing or unreadable, pull the client
name from any invoice/claim under:

```
/Projects/<Suburb> - <Address>/Progress Claim/PC#/
```

If neither is available, ask Paul directly.

> Worked example — 137 Annandale:
> `…/Tender Information/Write Up/R8.Annandale,137AnnandaleStreet.WriteUp.251027.docx`
> → Client: Jon and Stacey Day.

## Steps

1. **Confirm the project folder name** under `/Projects/` (e.g.
   `Annandale - 137 Annandale Street`).
2. **Read client details** from the Write Up (default path above). Confirm the
   client name(s) and email(s) with Paul if the Write Up doesn't carry email
   addresses.
3. **Get the contract date and original sum from Paul.**
4. **Generate the register:**

   ```bash
   python3 scripts/setup_project_register.py \
       --project-code "137a" \
       --project-name "137 Annandale Street, Annandale" \
       --client "Jon and Stacey Day" \
       --client-email "jonathon.day@example.com; stacey@example.com" \
       --contract-date 2025-08-20 \
       --contract-sum 2997386 \
       --dropbox-folder "Annandale - 137 Annandale Street" \
       --out "working/137a.Variation Register.xlsx"
   ```

5. **Upload** the generated `.xlsx` into the project's existing `Variations/`
   folder. (The Dropbox MCP can't write `.xlsx` directly, so the file is
   delivered to Paul to drop in, then verified by reading it back.)
6. **Verify** by reading the register back from Dropbox and confirming the
   Project Setup tab (client, contract date, sum, ABN `72 098 695 446`).

## Register filename convention

Match the project's existing file-naming style. For 137 Annandale the existing
files use the `137a.` prefix, so the register is:

```
137a.Variation Register.xlsx
```

One register per project, living alongside the `V#` folders inside
`Variations/`.

## Per-project layout overrides

Most projects file their `V#` folders directly under `Variations/`. Some don't,
and the registry in `scripts/config.json` records the exception via each
project's `variations_path`:

| Code | Project | `variations_path` | Next V# at last check |
|---|---|---|---|
| 137a | 137 Annandale Street, Annandale | `Variations` | V5 |
| 47t | 47 Tillock Street, Haberfield | `Variations` | V4 |
| 11e | 11 Eric Street, Lilyfield | `Variations` | V16 |
| 10w | 10 Woodlands Road, Taren Point | **`Variations/Phase 2`** | V13 |
| 44m | 44 Macintyre Crescent, Sylvania Waters | `Variations` | V39 |

**10w (Taren Point) is a staged job.** All variation work happens inside
`Variations/Phase 2/`. The `Variations/Phase 1/` section is historical and must
**never** be created in or modified. `create_variation.py --project-code 10w`
reads this from config and targets `Phase 2/` automatically; it also prints the
project note as a reminder. The 10w register therefore lives at
`…/Variations/Phase 2/`, not directly under `Variations/`.

(The "next V#" column is a snapshot — the live next number is always computed by
listing the relevant folder via the Dropbox MCP before allocating.)

## What never changes per project

- Builder identity on the notice: **Gold and Eagle Constructions trading as
  Renovate 8**, ABN **72 098 695 446**, NSW Contractor Licence **139701C**.
  These are baked into the notice template and the register's Project Setup
  tab defaults — they are not re-entered per project.
