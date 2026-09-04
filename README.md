# Balboa mapping agent — PyQt6 rewrite

## Setup

```
pip install -r requirements.txt
python main.py
```

First run creates `config.json`, `data/mappings_database.db`, and `data/backups/`
next to wherever you placed these files. Open **Settings** in the app to point
each fund at its actual cashbook file, and to point the database at your
shared folder if you want it there.

If the database lives on OneDrive/SharePoint, right-click that folder and
choose "Always keep on this device" so it's never a cloud-only placeholder.

## Layout

```
config.py                   absolute path resolution, fund settings
db.py                       the only module that opens a SQLite connection
mapping_engine.py           rapidfuzz matching, fund-scoped corrections
import_export_service.py    bank import, cashbook GL write-back, ETL export, archiving
legacy_import.py             one-time seed of transaction_history from a consolidated legacy file
workers.py                  QThread wrappers so long operations never freeze the UI
logging_setup.py            rotating file logger + uncaught-exception hook
ui/                         PyQt6 windows and dialogs — presentation only
main.py                     entry point
```

## Logs

Every run writes to `logs/balboa.log` (rotates at 2MB, keeps 5 backups) next
to wherever `config.json` lives. This replaces the original's `print()`-only
diagnostics, which would have gone nowhere the moment this is packaged as a
windowed executable. It also installs a `sys.excepthook`, so an unexpected
crash still leaves a stack trace in the log file even if nothing along the
way explicitly caught it. If something misbehaves in the field, that file is
the first thing to check.

Each layer only calls the one below it. The UI never touches SQLite or
openpyxl directly.

## What changed from the original script, and why

**Cross-fund correction bleed (fixed).** Corrections were keyed by
description text alone, so a GL correction made on one fund could get
silently applied to an identical-looking transaction on a different fund.
`corrections` is now keyed on `(fund_name, description)`, and the mapping
engine's fuzzy corpus is filtered to the current fund as well. Verified with
a direct test: an identical description on Fund I and Fund III now resolves
to two different GL codes.

**Same-day, same-amount transactions silently dropped (fixed).** The
cashbook-level duplicate check only compared `(date, amount)`, so two
legitimate transactions — a recurring transfer, a duplicate wire fee — could
collide and the second one would vanish with no warning. The dedupe key is
now `(date, amount, description)`. Verified directly: two rows with the same
date and amount but different descriptions are now both kept.

**Auto-archive was described but never wired up (fixed).** The original
import panel computed the cashbook's date range and logged it, but nothing
in the code actually archived a prior month once new data arrived.
`MonthArchiver.archive_completed_months()` now runs automatically after
every bank import: any month strictly before the one just imported, whose
rows are *all* already GL-mapped, gets archived to the database. A month
with any unmapped row is left alone and reported in the log instead of
being partially archived.

**Fuzzy matching swapped from fuzzywuzzy to rapidfuzz.** `fuzzywuzzy`
silently falls back to pure-Python `difflib` unless `python-Levenshtein` is
installed alongside it — and that package is GPLv2, worth avoiding on the
same trip where PyQt6 licensing was already a consideration. `rapidfuzz` is
MIT-licensed, C-accelerated, and `process.extractOne` runs the whole
candidate list in one call instead of a manual Python loop.

**Training data now reads live from the database.** The original tool kept
fuzzy-match training data in memory, populated only when someone manually
picked a "master training Excel file" — a step that had to be repeated
every session, and which silently produced empty results if forgotten.
The mapping engine now queries `transaction_history` directly, so a
correction saved this session, or a month archived a minute ago, is part of
the corpus for the very next match — no reload step, ever.

**Database path is no longer inferred from the working directory (fixed).**
The original used a bare relative filename for the SQLite file, so launching
the packaged app from a different folder silently created a fresh, empty
database. The path now comes from `config.json`, resolved from either the
`BALBOA_HOME` environment variable or the folder containing the running
executable — set once, explicit from then on.

**SQLite is defensive about network/cloud-sync storage.** Every connection
explicitly sets `journal_mode=DELETE` (WAL is documented as unreliable on
network filesystems) and a `busy_timeout`, so a sync client briefly locking
the file causes a short wait instead of an unhandled exception. `Database.backup()`
snapshots the file to `data/backups/` before any archive or import.

**Cashbook writes take a backup first.** Both the raw bank import
(`_append_to_cashbook`) and the GL write-back (`CashbookGLWriter`) copy the
live cashbook to a `.bak` file before writing, so a bad write during a
crash or a file-lock collision has a same-day rollback point.

**No more hardcoded personal paths in source code.** All fund cashbook
paths and the database location are edited from the Settings dialog and
persisted to `config.json` — nothing needs a code change to point at a
different machine or folder.

## Seeding from a consolidated legacy training file

If you already have a hand-built file combining years of prior mapped
transactions, use **Import legacy training data...** in the main window.
It's built specifically for this: pick the file with your own machine's
file dialog (nothing is sent anywhere), it previews before writing
anything, and it's safe to run twice — already-imported rows are
recognized and skipped, not duplicated.

It expects a sheet (default "All Funds") with columns matching, by name
rather than position: `JE comments`, `Yardi Account #`, `Fund`, and
`Amount` are required; `Yardi Account name` and `Post Date` are read if
present. Rows are de-duplicated on `(fund, description, GL code, amount)`
before anything is written — this matters if the file was assembled by
merging several older exports, since that tends to produce exact
back-to-back duplicate rows. Any `Fund` value that doesn't match a
configured fund name, or any row missing a GL code, is skipped and
called out in the preview rather than silently dropped or guessed at.

## Performance, at the volume discussed

Training corpus today: ~2,500 rows. At 30–90 new transactions/month across
the three funds over the next 7 months, the corpus grows to roughly
2,700–3,100 rows. Matching a month's batch against that corpus is on the
order of 80,000–280,000 rapidfuzz comparisons, which runs in well under two
seconds — and it now runs on a background thread (`MappingWorker`) with a
progress bar regardless, so the window stays responsive even as that number
grows over the coming years.

## Edge cases hardened for real-world use

- **GL write-back could misattribute two transactions that share a date
  and amount.** Fixed to also match on description, and to never reuse
  the same cashbook row for two different transactions.
- **Accounting-format negatives — `(1,234.56)`, common when a cell is
  stored as text — were silently excluded** as "not numeric" in the raw
  bank-import path, even though the manual-review path already handled
  them. Both paths now share one amount parser.
- **Two genuinely distinct transactions sharing an identical date,
  amount, and description within the same import batch** used to have
  the second one dropped as a false duplicate the instant the first was
  added. Fixed with a count-based check instead of a plain set — worth
  being precise about what this does and doesn't do: within one import
  run, duplicate-looking rows are trusted and both kept, since a bank
  file listing the same line twice almost always means two real
  transactions. Re-running the *same* statement in a later, separate
  session still correctly skips rows already on the sheet — that
  cross-run protection is the one that actually matters, and it still
  holds.
- **A bank account not yet in `sheet_mapping`** (a fund opening a new
  account, for instance) used to fail silently — the code fell back to a
  sheet name that didn't exist in the cashbook, and that account's entire
  month vanished with only a scrollable log line marking it. Now it's
  caught, the dollar total is computed, and it surfaces as a hard warning
  dialog naming the account and the amount that wasn't imported.
- **A blank or whitespace-only bank memo** could get fuzzy-matched
  against the corpus as if it were real text. Now treated the same as no
  description at all.

**One identified but deliberately not fixed yet:** the same description
text can mean opposite things depending on transaction sign — "WIRE
TRANSFER" could be a capital contribution coming in or a distribution
going out, same memo, opposite GL treatment. Matching and corrections
are currently text-only and don't consider whether the amount is
positive or negative. Making corrections and the fuzzy corpus sign-aware
is the right fix, but it changes matching behavior across the whole app —
worth planning deliberately, not rushing in.

## Import chains straight into the review table

"Import bank file..." no longer requires a separate "Upload monthly
file" step afterward re-reading the same file that was just imported.
Once an import finishes, the app automatically loads the just-imported
sheets into the review table — same cashbook write as always, same GL
mapping workflow, one fewer manual round trip. If some accounts in the
import weren't in `sheet_mapping` and got skipped (see the unmapped-account
warning above), those are excluded from the auto-load too — there's
nothing on disk for them to read yet.

## Editing an archived transaction

Double-click a row in the database viewer's transaction history tab (or
select it and click "Edit selected row") to fix a bad GL code, a
mis-mapped fund, or a typo from a legacy import. A backup is taken
before every edit, same as every other write in this app. Reassigning
the fund updates `property_code` automatically — there's no way to end
up with a row whose fund and property code disagree. Deleting a row now
also takes a backup first, which it previously didn't.

## Completed months no longer keep reappearing

Reading a cashbook sheet always reads the whole sheet — that hasn't
changed and isn't going to, since the cashbook is meant to hold every
month permanently. What changed: both `_upload_file` and the auto-chain
after "Import bank file..." now filter the loaded rows against
`Database.get_archived_keys()` before showing them in the review table.
A row already archived for that fund is dropped from the view, not from
the Excel file — it's still there, it just stops competing for
attention with what's actually new this cycle.

This depends entirely on the month having actually been archived. If
"Save GL to Cashbook" was skipped for a month, its rows never got GL
codes written back into the sheet, so the auto-archive check (which
reads those same columns) has no way to know the month is done — and it
will keep reappearing until that step is completed and the month is
archived, either automatically on the next import or manually via
"Archive reviewed batch."

## Starting a new month with a genuinely clean cashbook

"Start new month..." snapshots the current fund's cashbook to a dated
file in an `Archive/` folder next to it, then clears every data-row
column (not just the date column — clearing only column A would let a
new September transaction land on a row that still had August's GL code
sitting in columns F/G) so the live file is genuinely empty and ready
for the next import. The archive copy is written and verified to exist
*before* the live file is opened for writing — if the snapshot fails
for any reason, the live cashbook is never touched.

Before clearing, it checks which rows on the live sheet aren't yet in
the database for that fund and warns you by name and count. Those rows
are still fully preserved in the archive snapshot either way — the
warning is about them falling out of the *active* mapping/training
workflow, not about losing them.

This only makes sense because the cashbook doesn't carry any
running/cumulative balance across months and reconciliation only needs
the current month live — if either of those changes, this feature needs
revisiting before continuing to use it.




