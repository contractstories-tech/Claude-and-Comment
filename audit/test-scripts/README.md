# Audit reproduction scripts

Playwright/Node scripts used to execute the browser-app portion of `../AUDIT_REPORT.md`
against the supplied `Clause_Comments_Library.html`. Not part of the product — these are the
audit's test harness, kept for reproducibility.

## Requirements

- Node.js, with `playwright` resolvable (e.g. `npm install -g playwright`, or set
  `NODE_PATH` to a global install).
- A Chromium executable (adjust the `executablePath` in each script if yours differs).
- Copy the supplied `Clause_Comments_Library.html` into this directory as `app.html` before
  running any script — it is the audit subject and is intentionally not committed here.

## Scripts

| Script | What it exercises |
|---|---|
| `test1_basic.js` | Initial load, console/network capture, File System Access API availability on `file://` |
| `test2_full.js` | CRUD, favourites, duplicate, search, Unicode/XSS probe, long clause, JSON backup, Save As download, delete |
| `test3_lifecycle.js` | Round-trip reopen, crash-recovery reload, import merge/conflict, malformed/unrelated/foreign file import |
| `test4_recovery_root_cause.js` | Direct IndexedDB key inspection proving the libraryId-rotation root cause of Finding C1 |
| `test5_recovery_after_save.js` | Confirms recovery works correctly once a library has been saved at least once |
| `test6_perf.js` | Load/search/filter/detail/save timing at 500/2,000/5,000 synthetic entries |
| `test7_cache_size.js` | Plain-text vs WordOpenXML character-volume projection at scale |
| `test8_script_break.js` | `</script>`-breakout payload in clause text — confirms no code execution on save/reopen |
| `test9_misc.js` | Rich-XML fresh→stale transition, keyboard shortcuts |
| `test10_no_idb.js` | Recovery behaviour with `indexedDB` undefined (localStorage fallback) |
| `test11_save_button_no_fsaccess.js` | Two consecutive main-Save-button downloads on a browser without the File System Access API — reproduces Finding H1 |
| `gen_large.js` / `build_seeded_html.js` | Generate synthetic libraries and seed them into a copy of the app for the perf tests |

Run any script with `node <script>.js` after setting `NODE_PATH` to wherever `playwright` is
installed, e.g.:

```
NODE_PATH=/opt/node22/lib/node_modules node test2_full.js
```

Output files land in `./out/` (gitignored).
