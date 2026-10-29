# Changelog

## 2.1.0 — 2026-10-29
- Added: **Combine duplicate IDs** option (step 2, on by default). When the shared file has the same ID on several rows, all their non-blank values are joined with a separator you choose, repeated values are dropped, and it works with both Replace and Concat rules. Turn it off to get the previous behaviour (last row wins).
- Added: the preview reports how many IDs are duplicated in the shared file and how many extra rows that is.
- Added: the summary counts cells built from duplicates, and the change log marks each one ("Combined N values from duplicate IDs").
- Saved rule sets now include the combine setting and separator.
- Demo `shared.xlsx` now contains a duplicated ID.

## 2.0.0 — 2026-10-29
First public release (generalised from an internal requirements-matrix merge tool).

- **Fixed:** formatting was lost on save. The free SheetJS library cannot write cell styles, so fills, fonts and freeze panes disappeared even though the tool said "formatting preserved". The tool now patches only the changed cells inside the `.xlsx` package, so everything else is kept.
- **Fixed:** "Text length must not exceed 32767 characters" errors — the limit is now applied only to cells the tool writes, and untouched cells are never re-saved.
- **Fixed:** with duplicate keys in the original, only the last copy was updated; now all copies are.
- Added: offline bundle (no internet needed), header row selection, ignore-case matching, auto-map same-named columns, save/load rule sets, CSV change log, recalculation on open, `.xlsm` support, drag-and-drop, dark mode, mobile layout.
- Added: automated test-suite, demo workbooks, PDF user guide.

## 1.x — March 2026
Internal versions: key matching, Replace/Concat rules, blank-value retention, 32,767-character guard.
