# Nepali Date Converter (AD → BS, Excel)

## Intent

Convert English (Gregorian/AD) dates to Nepali (Bikram Sambat/BS) dates in Excel.
Goal: select one or more cells containing AD dates, run a macro, get the BS
equivalent — without depending on a live API and without losing accuracy.

Use case is narrow: converting past dates, roughly 1-2 years back from whenever
this is run. Table below covers far more than that on purpose (see "Why the
full table" below).

## Data source

[Project_Parva](https://github.com/dantwoashim/Project_Parva)
(`backend/app/calendar/constants.py`), pulled directly from the raw GitHub
file — not from an AI-summarized fetch. Two earlier attempts to summarize the
repo via WebFetch gave contradictory/wrong answers about year range and epoch,
so don't trust an LLM's paraphrase of this kind of data again; always pull the
raw file and read it directly.

- Coverage: BS 2000-2099 (~AD 1943-2043), one static lookup table
- Anchor/reference point: BS 2082-01-01 = AD 2025-04-14 (per Hamro Patro)
- Conversion method: sum month lengths forward/backward from the anchor to
  find the target date. Same approach the source project uses.

## Confidence tiers (important)

Project_Parva's own `provenance.py` grades the table — it is **not** uniformly
verified:

- **BS 2078-2083** → `official` — backed by structured official source data
- **BS 2076-2077** → archived official PDF exists, not yet structured/verified
- **Everything else (2000-2099)** → `static_table_unverified` — in the table,
  but the project itself won't vouch for it

The macro appends a literal `[unverified]` flag to any output outside
2078-2083. This is not automatic/magic — it's a hardcoded range check.

Practically: since actual usage is "1-2 years back," current dates land inside
2082-2083 — the verified core. The wider table is there so the tool doesn't
need touching again for years, not because those years are equally trustworthy.

## Why the full table (size concern, resolved)

The full 100-year table is a few KB (100 years x 12 month-lengths, each ≤32,
fits in a byte). Negligible size difference vs. only storing 1-2 years, so no
reason to trim it.

## Files

- `AD_to_BS_Converter.bas` — VBA module for **Microsoft Excel** (Windows/desktop).
  Import via Alt+F11 → File → Import File. Run `ConvertSelectionToBS` on a
  selection of AD-date cells; result **overwrites the selected cell in
  place**, forced to Text format. Not reversible except via Ctrl+Z (which
  only survives until your next action) — keep a backup column/copy of the
  AD dates first if you need to keep both versions.

  **Result summary**: shows a native `MsgBox` (converted count, skipped
  count, native skipped row numbers) — this works fine in real Excel, unlike
  ONLYOFFICE. For consistency with the ONLYOFFICE port, the same summary is
  also written two rows below the selection (not a fixed cell like A1) and
  selected.

- `AD_to_BS_Converter.onlyoffice.js` — JavaScript port for **ONLYOFFICE Desktop
  Editors** (tested on Community 9.4.0.129, Linux). Access via **View tab →
  Macros** (no Developer tab in ONLYOFFICE, unlike Excel), paste the whole
  file in, select AD-date cells first, run. Same overwrite-in-place behaviour
  as the VBA version.

  **Result summary**: no native dialog is available or usable here.
  `alert()`/`window.alert()` has been blocked in ONLYOFFICE macros since
  v7.1 (confirmed via their community forum). `confirm()` was tried as an
  alternative and **confirmed not to work** on 9.4.0.129 (tested — no dialog
  appeared, and it doesn't throw an exception either, so a script can't even
  detect the failure to fall back automatically). Dropped entirely. Instead,
  the macro always writes the summary (converted count, skipped count,
  native skipped row numbers) two rows below your selection — not a fixed
  cell like A1, to avoid overwriting unrelated data — and selects it.
  ONLYOFFICE has no reliable "scroll into view" API (known limitation, no
  fix as of this writing), so that selection may not visibly scroll on
  screen — check the Name Box if you don't see the highlight move.

  **Two assumptions in this file are unverified against the real app** (the
  official API docs didn't state them explicitly) — check these first if
  output looks wrong:
  1. `GetValue()` on a date-formatted cell returns a numeric Excel-style
     serial number, not a string or JS Date object. If wrong, every date
     cell will silently count as "skipped."
  2. `GetRow()` is 0-based. If wrong, reported row numbers (skipped rows,
     fallback summary row) will be off by one.
  Check against the known-good test dates below before trusting it on real
  data.

## Compatibility notes

- **Excel (Windows)**: works as-is (standard VBA). Not yet tested — pending
  Windows access.
- **Excel Online (Microsoft 365 web)**: VBA does not run in the browser editor
  at all. This macro cannot work there. Office Scripts (TypeScript) would be
  the equivalent if that's ever needed.
- **OnlyOffice Desktop Editors** (tested environment: Community 9.4.0.129,
  Linux): macros are **JavaScript**, not VBA — accessed via **View tab →
  Macros** (no Developer tab, unlike Excel). Ported: see
  `AD_to_BS_Converter.onlyoffice.js` above. Not yet run in the actual app —
  see the two unverified assumptions noted there.

## Date-format issue (separate from BS conversion)

Two distinct problems, don't conflate them:

1. **BS output isn't a real Excel date type** — Bikram Sambat has no native
   Excel date type, so output is always text. The macro forces the target
   cell to Text format before writing (`NumberFormat = "@"`), rather than the
   "insert a leading space" hack some users rely on.
2. **AD dates displaying as M-D-Y instead of Y-M-D** — this is Windows
   regional locale bleeding into Excel's default short-date display, unrelated
   to BS conversion. Fix without a system-wide change: Format Cells → Custom
   → `yyyy-mm-dd` (per cell/workbook), or edit the "Normal" cell style /
   `Book.xltx` template to make it the default for new cells. Changing the
   Windows regional short-date setting instead would affect every app on the
   machine, not just Excel — only do that if that's actually wanted.

## Known-good test dates (Nepali New Year / Baishakh 1)

Output now overwrites the cell in place, so test on a throwaway copy of these
dates (not real data) — paste into a column, run the macro, compare against
the expected values below:

```
2020-04-13   → expect 2077-01-01   (tests the [unverified] flag)
2022-04-14   → expect 2079-01-01
2023-04-14   → expect 2080-01-01
2024-04-13   → expect 2081-01-01
2025-04-14   → expect 2082-01-01   (anchor date)
2026-04-14   → expect 2083-01-01
```

## Status / open items

- [ ] Test macro in real Excel on Windows
- [ ] Test `.onlyoffice.js` in ONLYOFFICE Desktop Editors on Linux — confirm
      the two unverified API assumptions noted above
- [ ] Not tested against an independent third-party BS calendar — the test
      dates above are public/well-known reference points, but the macro's
      full-range output (outside those specific dates) hasn't been
      cross-checked against anything but Project_Parva's own table.
