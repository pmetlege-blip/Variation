# Variation Numbering Scheme

## Format

```
V[n]            (e.g. V1, V2, V10, V11)
```

Matches the existing Renovate 8 convention in Dropbox — looking at
`/Projects/Annandale - 137 Annandale Street/Variations/`:

- V1
- V2 - Asbestos under Concrete
- V3 - Absestos Removal
- V4 - Bamboo Works and water detailing

The next allocated variation in that project would be `V5`.

## Rules

1. **Per-project sequential.** Reset to `V1` for each new project.
2. **No zero-padding.** `V10`, not `V010`. (This matches the existing
   convention; if any project hits V100+, sorting still works because the
   register column sorts numerically, not alphabetically.)
3. **No gaps.** Cancelled variations keep their number. Move the folder to
   `Variations/Archive/` if you want it out of the way; the register row
   stays with status `CANCELLED`.
4. **No sub-numbers.** No `V7a`, no `V7.1`. If a variation needs amending
   after issue, either:
   - reissue under the same V# with a new notice version (`v2`, `v3`, ...),
     OR
   - raise a new V# if scope materially changes.

## Where the number lives

- **Folder name:** `V5 - extra power points kitchen/`
- **Notice header:** `Variation Number: V5`
- **Notice filename:** `V5_v1_RAISED.pdf`
- **Register row:** column A, value `V5`
- **Progress claim / invoice line:** `V5 — [short description]`

## How the next number is allocated

When Claude allocates a new variation for a project:

1. List the project's `Variations/` folder via Dropbox MCP.
2. Find all entries matching `V<n>` or `V<n> - ...`.
3. Take `max(n) + 1`.
4. (Cross-check) Read the register's VAR # column and take the max from
   there too. The higher of the two wins, so cancelled-but-not-yet-archived
   variations aren't reused.

If you're migrating mid-project and want to force a specific number, use
`scripts/create_variation.py --number 7`.

## Versioning within a single V#

Inside one V#, the notice goes through versions as the variation moves
through phases:

- `V5_v1_RAISED.pdf` — Phase 1 (price TBC or indicative).
- `V5_v2_FINALISED.pdf` — Phase 2 (confirmed price).
- `V5_v3_AMENDED.pdf` — rare; only if a finalised notice needs correcting.

Each version is a new file in `03_Notice/` — never overwrite a prior
version, so the audit trail is intact.
