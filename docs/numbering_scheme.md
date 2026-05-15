# Variation Numbering Scheme

## Format

```
[PROJECT_CODE]-VAR-[###]
```

Examples:
- `42SMITH-VAR-001`
- `42SMITH-VAR-002`
- `17BRIDG-VAR-001`

## Rules

1. **Sequential per project.** Numbers start at `001` for each new project
   and increment by 1. There is no global cross-project counter.
2. **Zero-padded to 3 digits.** Allows clean sorting in Dropbox and Excel.
   If a project somehow exceeds 999 variations, expand to 4 digits and
   reissue the master register.
3. **No gaps.** Even cancelled variations keep their number — the row in
   the register is marked `CANCELLED` but the number is never reused.
4. **No sub-numbers** (no `VAR-007a`, no `VAR-007.1`). If a variation needs
   amending after it's been issued, either:
   - reissue the same number with a new version (`v2`, `v3`, ...) — for
     a Phase 1 → Phase 2 transition, or for correcting a typo, OR
   - raise a new variation number — if the scope materially changes.
5. **Cancellations don't reset the counter.** If `VAR-005` is cancelled,
   the next variation is still `VAR-006`.

## Where the number lives

- **Folder name**: `42SMITH-VAR-007 - extra-power-points-kitchen/`
- **Notice header**: `Variation Number: 42SMITH-VAR-007`
- **Notice filename**: `VAR-007_v1_RAISED.pdf`
- **Register row**: column A, value `42SMITH-VAR-007`
- **Invoice / progress claim reference**: `VAR-007`

## How to allocate

Manually:

1. Open the project's master register.
2. Find the highest existing variation number for that project.
3. Add 1, zero-pad.

Via script:

```bash
python3 scripts/create_variation.py --project 42SMITH --description "..."
```

The script reads the register, allocates the next number, creates the
folder, and stubs an initial row. (It does NOT generate the notice — that
is still done from the template.)

## Versioning within a variation

Within a single variation number, the notice goes through versions:

- `v1_RAISED` — Phase 1 notice, price TBC or indicative.
- `v2_FINALISED` — Phase 2 notice, confirmed price.
- `v3_AMENDED` — only if the FINALISED notice needs correcting (e.g.,
  client-requested scope tweak, typo, recalculation). Rare.

Each version is a new PDF in `03_Notice/`. **Never overwrite** an older
version — keep the history for the audit trail.
