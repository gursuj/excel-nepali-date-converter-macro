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
  selection of AD-date cells; each selected cell is overwritten **in place**
  with its BS equivalent, forced to Text format. The original AD date is gone
  once the macro runs — **running any macro clears Excel's entire undo
  history**, so Ctrl+Z will not bring it back (confirmed via real testing,
  not just theory — this is stronger than "not reversible except Ctrl+Z").
  Keep a copy of the AD dates elsewhere first if you need to preserve them.

  **Result summary**: shows a native `MsgBox` (converted / already-BS /
  skipped counts, native skipped row numbers, a sample raw value+type for
  the first genuinely unparseable cell) — this works fine in real Excel,
  unlike ONLYOFFICE. For consistency with the ONLYOFFICE port, the same
  summary is also written two rows below the selection (not a fixed cell
  like A1) and selected. Columns are also auto-fit after conversion so long
  values (e.g. the `[unverified]` suffix) aren't clipped.

  **Bug found and fixed (2026-09-16)**: a real dataset (CSV-origin data,
  `ENROLLED DATE` column) came back as "Converted: 0, Skipped: 273" even
  though most of column C already showed correctly converted BS dates. Root
  cause: the macro had been run once successfully, then run again on the
  same range — and it counted the already-converted `yyyy-mm-dd` text as
  "not a date" instead of recognizing it, so every cell (converted or not)
  got flagged as skipped on the second pass. Fixed by:
  - Recognizing already-BS-formatted text (`yyyy-mm-dd`, optionally
    `[unverified]`) up front and treating it as "already converted, skip
    safely" — re-running the macro on the same range twice is now safe.
  - Falling back to `.Text` (and a trimmed string) if `.Value` doesn't parse
    as a date, since CSV-origin data can end up as plain text rather than a
    real date value.
  - Reporting a sample skipped cell's raw content and type in the summary,
    so a genuine parse failure is self-diagnosing without needing to
    reproduce it in Excel directly.

## Installing / removing the module (Excel, Windows)

Two ways to reach the VBA editor: the keybind (Alt+F11, works regardless of
ribbon setup) or the Developer tab (Developer → Visual Basic). The Developer
tab isn't shown by default — enable it once via File → Options → Customize
Ribbon → tick "Developer" in the right-hand list → OK.

**Add:**
1. Open the workbook. Reach the VBA editor via Alt+F11, or Developer →
   Visual Basic.
2. File → Import File (menu only — no keybind for this).
3. Select `AD_to_BS_Converter.bas`.
4. Save the workbook as a macro-enabled file (`.xlsm`), otherwise the macro
   won't be there next time you open it.

**Remove:**
1. In the VBA editor, click the `AD_to_BS_Converter` module in the Project
   Explorer to select it.
2. File → Remove AD_to_BS_Converter (there's no right-click "Remove" option
   for modules in the classic VBA editor, and no keybind — it has to go
   through the File menu).
3. Choose "No" if it asks whether to export first (unless you want to keep a
   copy).
4. Save the workbook.

**Does it persist?**
The module is saved inside whatever workbook you imported it into — it's
part of that file, not a global Excel setting. Close and reopen that same
workbook and the macro is still there. Open a *different* workbook and it
won't have the macro unless you import it into that one too.

If you want it available in every workbook without re-importing each time,
put it in `PERSONAL.XLSB` (Excel's hidden personal macro workbook, created
via View → Macros → Record Macro → "Store macro in: Personal Macro
Workbook", or by importing directly into it if it already exists) — that
file loads automatically every time Excel starts, regardless of which
workbook you open.

## Running the macro

1. Select the cell(s) containing AD dates.
2. Run `ConvertSelectionToBS` one of two ways:
   - Keybind: Alt+F8 → select `ConvertSelectionToBS` from the list → Run.
   - GUI: Developer → Macros (same Alt+F8 dialog, reached via the ribbon
     instead) → select `ConvertSelectionToBS` → Run.
3. A popup reports converted / already-BS / skipped counts, native skipped
   row numbers, and (if anything was genuinely unparseable) a sample raw
   value to help diagnose it. The same summary is also written two rows
   below your selection and selected.

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
  the macro always writes the summary (converted / already-BS / skipped
  counts, native skipped row numbers, a sample raw value+type for the first
  unparseable cell) two rows below your selection — not a fixed cell like
  A1, to avoid overwriting unrelated data — and selects it. ONLYOFFICE has
  no reliable "scroll into view" API (known limitation, no fix as of this
  writing), so that selection may not visibly scroll on screen — check the
  Name Box if you don't see the highlight move.

  **Handles both real date values and plain text dates** (CSV-origin data,
  `yyyy-mm-dd`, or `D-Mon-YY`/`D-Mon-YYYY` like `3-Apr-26`) — see the Excel
  bug entry above; the same root cause applied here since the port
  originally only handled numeric serials. Already-BS-formatted text is
  recognized and left alone, so re-running on the same range twice is safe.

  **One assumption in this file is unverified against the real app** (the
  official API docs didn't state it explicitly) — check this first if
  reported row numbers look off by one: `GetRow()` is 0-based.
  Check against the known-good test dates below before trusting it on real
  data.

## Compatibility notes

- **Excel (Windows)**: tested against a real dataset — found and fixed both
  a VBA syntax error (the original giant `Array(...)` table literal exceeded
  VBA's line-continuation limit; rewritten as one `SetRow` call per year
  instead) and the false-skip bug documented above.
- **Excel Online (Microsoft 365 web)**: VBA does not run in the browser editor
  at all. This macro cannot work there. Office Scripts (TypeScript) would be
  the equivalent if that's ever needed.
- **OnlyOffice Desktop Editors** (tested environment: Community 9.4.0.129,
  Linux): macros are **JavaScript**, not VBA — accessed via **View tab →
  Macros** (no Developer tab, unlike Excel). Ported: see
  `AD_to_BS_Converter.onlyoffice.js` above. `confirm()` confirmed not to
  work as a dialog (see above); one indexing assumption (`GetRow()` 0-based)
  still unverified.

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

- [x] Test macro in real Excel on Windows — done via real dataset (CSV
      import), found and fixed the "everything skipped" bug (see above)
- [ ] Test `.onlyoffice.js` in ONLYOFFICE Desktop Editors — `confirm()`
      dialog confirmed not to work (fixed, see above); the already-BS /
      text-date handling fix hasn't been run there yet; one indexing
      assumption (`GetRow()` 0-based) still unverified
- [ ] Not tested against an independent third-party BS calendar — the test
      dates above are public/well-known reference points, but the macro's
      full-range output (outside those specific dates) hasn't been
      cross-checked against anything but Project_Parva's own table.
