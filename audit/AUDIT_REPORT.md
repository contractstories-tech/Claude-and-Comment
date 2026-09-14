# Clause & Comments Library v6.1 — Independent Audit

Full formatted report (recommended): the published artifact linked in the session that
produced this commit. This file is the durable, plain-text record of the same findings.

Scope: `Clause_Comments_Library.html` (the browser app), `ClauseLibrary_Word.bas` (the Word
VBA module), `ClauseLibrary_FormCode.txt` / `ClauseLibrary_SearchFormCode.txt` (optional
UserForm code-behind), and `Word_Search_Workbench_Setup.txt` (setup notes) — the five files
supplied for review.

No Microsoft Word installation was available in this environment. The browser app was
executed for real in Chromium via Playwright, driven headlessly against the exact supplied
HTML file over `file://`. Every claim below is labelled **[Executed]**, **[Source-reviewed]**,
or **[Requires real Word testing]** — nothing executed is reported as merely inferred, and
nothing inferred is reported as tested.

Reproduction scripts for every executed test live in `audit/test-scripts/` in this repo.

---

## 1. How the system actually works

Five loosely coupled local files, no server:

- **The HTML master** (`Clause_Comments_Library.html`) — a single self-contained file. One
  `<script type="application/json" id="library-data">` tag holds the entire repository
  (entries, revision counter, `inboxSeenIds`); a vanilla-JS IIFE renders the UI from that
  JSON and rewrites the whole file on save. **This embedded JSON is the single source of
  truth.** IndexedDB, the Word cache, and the inbox are all caches or staging areas for it,
  never authoritative themselves.
- **ClauseLibrary_Word.bas** — the Normal.dotm module. Adds two right-click commands and two
  keyboard shortcuts (`Ctrl+Alt+Shift+C` capture, `Ctrl+Alt+Shift+I` search/insert) via
  `AutoExec`.
- **The inbox** (`%Documents%\ClauseLibrary\clause_inbox.jsonl`) — append-only, one JSON
  object per line, written by Word, polled by the browser every 6 seconds via the File
  System Access API.
- **The Word cache** (`...\clause_library_cache.cclcache`) — a tab-delimited snapshot the
  browser regenerates on every successful Save. The only thing Word ever reads to search or
  insert. Fully disposable.
- **The optional search UserForm** (`frmClauseLibrarySearch`) — a richer modeless
  search/insert workbench. `SearchFormReady` checks every required control exists and
  silently falls back to the compact `InputBox` flow if the form is missing or incomplete.

**Word → Library (capture):** `CaptureSelection` reads `Selection.Range`. If the range has
zero tracked-change revisions it takes `rng.Text`/`rng.WordOpenXML` directly; if it has
revisions, it builds a hidden temp document, inserts the XML (or falls back to
`FormattedText`), calls `AcceptAllRevisions`, and re-extracts clean text/XML — avoiding
baking tracked-change markup into a reusable precedent. `AppendEntryToQueue` serializes the
result to one JSON line, appended to the inbox with a retry loop against file locks.

**Library → Word (retrieval):** Word never reads the HTML file, only the `.cclcache`
snapshot. `InsertFromClauseLibrary` parses it, scores matches against title/category/
tags/wording/comments/context/source, and on insertion tries `Range.InsertXML` with the
stored base64 WordOpenXML first; a thrown error or a stale rich flag prompts before falling
back to plain text. Negotiation comments only ever attach via `Comments.Add` against a
wording selection the user made — a negotiation note cannot land in the contract body
silently.

**Safety nets — three independent layers:**

| Mechanism | Protects against | Durability |
|---|---|---|
| `clause_inbox.jsonl` | Browser closing/crashing after Word captured something, before ingestion | Durable — plain append-only file, replayed idempotently against `inboxSeenIds` on every reconnect |
| IndexedDB/localStorage recovery | Browser closing/crashing after direct-UI edits | Debounced (350ms) snapshot keyed by `libraryId` — **see Finding C1: broken for a never-yet-saved library** |
| `previousBackup` embedded in the HTML | A bad overwrite on next Save | One generation deep, restorable via "Restore previous" |

**When a piece is unavailable:** no File System Access API → every file op degrades to a
browser download, cleanly, no crash (but see H1/H3). No IndexedDB → falls back to
localStorage for recovery only **[Executed — confirmed clean degradation]**. Cache
missing/corrupt/wrong-version → specific correct message, refuses to proceed. Inbox write
blocked → retried 4× (~450ms), then a clear non-silent failure dialog.

---

## 2. Execution testing — what actually ran

Real Chromium via Playwright, `file://` load of the exact supplied HTML. Two notes before the
results: in this Chromium build, `file://` is treated as a secure context, so
`showOpenFilePicker`/`showSaveFilePicker` are available with no local server needed —
contrary to a common assumption. Headless automation cannot drive the native OS picker those
APIs open, so picker-based flows (Save-in-place, Connect Inbox, Connect Cache) were exercised
with the API deliberately removed to force the documented no-API fallback — which is also
exactly the path every Firefox/Safari user hits on *every* save, so it got the most scrutiny.

| Test | Method | Result |
|---|---|---|
| Initial load — console/network | Executed | Clean: 0 console errors, 0 external requests, `isSecureContext:true` on file:// |
| Create/edit/duplicate/delete entries | Executed | All work; duplicate resets favourite, forces Needs review |
| Favourites, tags, status/position/category filters, search | Executed | All compose correctly, verified against 500–5,000-entry synthetic libraries |
| Unicode/emoji/CJK/quotes/em-dash in clause text | Executed | Byte-for-byte round trip through save→reopen |
| HTML/script-like strings in title & clause text (XSS probe) | Executed | Rendered inert everywhere; a `</script><script>` break-out payload survived save/reopen without executing |
| 84,000-character clause | Executed | No lag, no truncation |
| JSON Backup export | Executed | Valid, complete, re-parses cleanly |
| Save (main button), no File System Access API | Executed | **Finding H1** — two consecutive saves produced two files both stamped r1 with different content; "unsaved" indicator never cleared |
| "Save As"/Download Copy, no File System Access API | Executed | **Finding H3** — revision/savedAt not stamped on this path |
| Reopening a saved library (round trip) | Executed | All entries, revision, Unicode intact |
| Crash recovery, edits made, reload, never-yet-saved library | Executed | **Finding C1 — all unsaved work lost**, silently |
| Crash recovery, same, on a library saved at least once | Executed | Correct — "Recovered unsaved changes", 4th entry reappeared |
| Crash recovery with IndexedDB disabled (localStorage fallback) | Executed | Same C1 failure — root cause independent of storage backend |
| Import: merge with same-ID conflict | Executed | Correct dialog, correct count, prompt()-based resolution behaves as documented |
| Import: malformed JSON / unrelated JSON / foreign HTML | Executed | Each rejected with a specific correct error; app stayed usable every time |
| Rich-text freshness (fresh→stale on clause edit) | Executed | Banner/pill flip correctly the instant clause text changes |
| Keyboard shortcuts (Ctrl+S, Ctrl+F, Ctrl+Shift+N) | Executed | All fire correctly |
| Synthetic libraries at 500/2,000/5,000 entries (3–23 MB) | Executed | See §8 below; 0 console errors at any size |
| Word capture/tracked changes/protection/InsertXML/corporate GPO/OneDrive/multi-instance | Source-reviewed only | No Word install available — see §3 |

---

## 3. Word / VBA integration audit

No Word installation was available. Everything below is **[Requires real Word testing]**
unless marked otherwise.

**Well built [Source-reviewed]:**
- Idempotent UI registration — `SetupContextMenu` removes its own tagged buttons before
  re-adding, and every `CommandBarButton` is `Temporary:=True`, so nothing persists into the
  CommandBars customization file and repeated `AutoExec` runs cannot accumulate duplicates.
- Revision-safe capture via a hidden temp doc + `AcceptAllRevisions`.
- The InsertXML fallback chain is correctly defensive: rich insert attempted first, thrown
  error or stale flag routes to an explicit prompt before falling back to plain text, wrapped
  in an `UndoRecord` custom record.
- Negotiation comments structurally cannot leak into contract body text (`UseCommentText`
  hard-requires an active selection).

**WordOpenXML round-trip — the core technical bet [Requires real Word testing]:**
Capturing `Range.WordOpenXML` and reinserting via `Range.InsertXML` into an unrelated
document is Microsoft's own documented technique, not a hack. Two structural limitations are
inherent to it, not bugs in this code: (1) a captured range referencing a bookmark/cross-
reference/field pointing *outside* the range carries a dangling reference once transplanted;
(2) list/numbering continuity is the most commonly reported real-world `InsertXML` failure —
a captured numbered clause can renumber unpredictably, and critically, a numbering glitch
that succeeds without throwing would **not** trigger this code's plain-text fallback, since
that fallback only fires on a thrown error. This is the single most important thing to verify
in real Word before relying on this for numbered/defined-term clauses.

**Deployment in a managed Microsoft 365 environment [Requires real Word testing]:** everything
lives in Normal.dotm via ordinary macro execution — no "Trust access to the VBA project
object model" dependency. Needs macros enabled at all (any policy blocking Normal.dotm
macros disables the tool silently — no menu items appear, with no explanation). `ADODB.Stream`
and `MSXML2.DOMDocument.6.0` are long-standing, near-universal Windows COM components. 32 vs
64-bit Office is irrelevant (the only pointer-sized API, `GetSystemTime`, is correctly
`#If VBA7`-guarded). OneDrive-redirected Documents folders should work transparently
(`Options.DefaultFilePath(wdDocumentsPath)`, not a hardcoded path), but OneDrive background
sync racing an open Word/browser handle on the bridge files is untested.

---

## 4. The inbox/cache bridge

Append-only JSONL inbox (durable one-way, Word→browser) + disposable tab-delimited cache
(browser→Word, regenerated every save). Neither is authoritative — the HTML master is —
which is what makes most of the concurrency questions resolve favourably:

- **Multiple Word instances writing simultaneously:** each write is one already-built line
  via `Open ... For Binary Access Write Lock Write`, retried 4× on contention. Should
  serialize correctly for realistic near-simultaneous captures; true millisecond-simultaneous
  writes were not tested (no second Word process available) — **[Requires real multi-instance
  Word testing]**.
- **Partial/interrupted writes:** the browser's line reader buffers any trailing partial line
  until a terminator appears — a write caught mid-flush is simply not read yet, not
  misparsed.
- **Stale inbox entries / deliberate deletion:** `inboxSeenIds` is checked independently of
  whether the entry still exists in `data.entries`, so deleting an entry does not resurrect it
  from a full inbox rescan.
- **Wrong-library cache:** checked both directions — the browser refuses to connect a
  mismatched `.cclcache`, and Word's `EnsureActiveLibrary` refuses to search a cache belonging
  to a different pinned-active library, with an explicit `ActivateCurrentClauseLibrary`
  escape hatch rather than a silent switch.

**Would a heavier architecture help?** No. A real embedded database or sync server would
solve problems this design doesn't have (realistically one browser tab and one interactive
Word session per person), while adding install/failure surface a "no IT ticket required" tool
shouldn't carry. Not recommended.

**Not verified:** OneDrive sync racing an open file handle; truly simultaneous multi-Word-
instance writes; behaviour when the bridge folder is renamed/moved out from under an
already-open browser tab.

---

## 5. Data safety & the entry lifecycle

| Crash point | Outcome |
|---|---|
| Word crashes after capture, before browser ingests | Safe — inbox line already durable |
| Browser closes after ingest, before Save, already-saved library | Safe — recovery restores it **[Executed]** |
| Browser closes after ingest, before Save, never-saved library | **Lost — Finding C1** |
| Browser closes after direct-UI edits, never-saved library | **Lost — same root cause, and the more common case on day one** |
| Save fails partway (disk full, permission denied) | Safe — error surfaced, `dirty` state untouched |
| Save "succeeds" via download-only fallback | Not lost, but bookkeeping is unreliable — **H1/H3** |
| Importing a stale/older backup over a newer library | Safe — Merge is content-diff-based, Replace is a separate confirmed action |

**Direct answer:** once a library has been saved to disk at least once, the recovery, backup,
and import machinery held up against a full day of adversarial testing — no silent loss or
corruption of already-persisted data was found anywhere. The one real gap is the window
between first opening the tool and the first Save, which currently is not protected at all.

---

## 6. Data model & playbook design

The taxonomy (Clause/Comment/Clause+Comment; Preferred/Fallback/Retired/Unclassified;
provider/balanced/customer/none position; category; tags; favourite; source; needs-review) is
a sensible first cut. Two gaps show up specifically at hundreds-to-thousands scale, not in a
small demo:

1. **No first-class relationship between variants of the same clause.** "Limitation of
   Liability — Preferred" and its Fallback sibling are connected only by title convention,
   category, and tags. At scale this becomes the library's biggest maintenance burden —
   nothing flags that editing one leaves siblings inconsistent, nothing prevents 5 near-
   duplicates accumulating over a year of ad-hoc capture.
2. **Status and Position are single-valued**, but real precedent varies by counterparty
   leverage, deal size, and jurisdiction simultaneously — a realistic path to a lawyer no
   longer trusting the classification enough to maintain it.

Neither is a defect in what shipped; both are worth deciding on deliberately before the
library reaches a size where restructuring is painful.

---

## 7. Usability as a lawyer's daily tool

**Works well:** capture friction is low (select → shortcut/right-click → short form → done).
Retrieval-side safeguards are the strongest part of the product: Retired entries get an
unmissable warning on every surface plus a confirmation before Copy/Insert; Needs-review
entries are excluded from default search and clearly labelled where they do appear; the
rich/stale indicator removes an entire class of "did this insert actually carry the right
formatting?" doubt; search results show status/position/category/source at a glance.

**Will grate:** the default Word-side search (without the optional workbench form) is an
`InputBox`/numbered-`MsgBox` loop — a real step down from the browser's instant filtered
list, and it's the *default* unless someone manually builds the optional UserForm. Import
conflict resolution is one blocking `prompt()` per conflict with no "apply to all" — fine for
1-2 conflicts, painful merging libraries that have diverged by dozens of entries. The
classification burden from §6 is a UX cost as much as a data-model one — nothing nudges
dedup/linking as the library grows.

---

## 8. Performance & scale

Synthetic libraries generated and loaded for real (20-30% of entries carrying simulated
WordOpenXML payloads):

| Entries | File size | Load | Search | Filter | Detail render | Save (download) |
|---|---|---|---|---|---|---|
| 500 | 2.97 MB | 483 ms | 244 ms | 74 ms | 159 ms | 331 ms |
| 2,000 | 12.0 MB | 1,339 ms | 250 ms | 99 ms | 613 ms | 1,211 ms |
| 5,000 | 23.4 MB | 1,703 ms | 253 ms | 200 ms | 843 ms | 2,229 ms |

Even at a moderate 20-30% rich-entry rate, XML text volume already exceeded plain-text
volume in every sample (e.g. at 2,000 entries: ~3.7M plain chars vs ~7.0M XML chars). Real
Word captures of short clauses routinely produce XML far larger relative to plain text than
this synthetic sample. Zero console errors at any tested size; a single-lawyer working
library (low hundreds to low thousands) will feel instant. A shared firm-wide library with
heavy rich-formatting usage in the high thousands is where cache size and VBA-side parse time
(untested — no Word available) merit watching.

---

## 9. Security & privacy

The offline claim holds **[Executed]**: zero external network requests observed on load;
source-wide grep confirms no CDN/font/telemetry/remote-script reference anywhere. No stored
entry field is ever written via `innerHTML` or any HTML-injection-capable sink — confirmed
both by grep and by two direct adversarial probes: a `<script>alert(1)</script>` title
rendered as inert text everywhere, and a `</script><script>window.__pwned=true</script>`
payload in clause text — specifically designed to break out of the embedded JSON `<script>`
tag on save — did not execute after a full save→reopen round trip (the app's own
`.replace(/<\//g,'<\\/')` escaping handled it correctly).

Malformed/adversarial imports (invalid JSON, unrelated-but-valid JSON, foreign HTML with no
library-data tag) were all rejected with specific correct messages; app stayed fully usable
after each. On the VBA side, the main practical exposure is the ordinary macro-security
posture of any Normal.dotm macro — runs with the user's own filesystem permissions inside
their Documents folder only; no path in the code accepts or interpolates externally
controlled paths.

---

## 10. Code quality

The JavaScript is dense but consistent; every normalization path (`normalizeEntry`,
`normalizePayload`) is defensive about missing/malformed fields and held up under every
adversarial import tested. The VBA is unusually disciplined for a module this size:
consistent error-handling pattern throughout, no dead code found, and the cache format's
escape/unescape pair is correctly ordered (backslash escaped first, traced by hand) so it
round-trips unambiguously. Two genuine pushbacks: the cache-format version check is a bare
hardcoded string comparison with no migration path; and reliance on native
`confirm()`/`prompt()` for import-conflict resolution is a maintainability smell as much as a
UX one — synchronous, untestable headlessly by anything but a real browser, doesn't scale
past a handful of conflicts.

---

## 11. Findings, ranked

### CRITICAL

**C1 — Unsaved work on a never-yet-saved library is lost on crash/reload, with no warning.**
The distributed template's embedded JSON always carries `libraryId:"lib-initial"`. On every
fresh load, `init()` replaces it with a brand-new random ID *before* checking for a recovery
snapshot. The 350ms-debounced recovery snapshot is keyed by that ID. Reload the same
never-saved file and a second random ID is minted — the recovery lookup under the new ID
finds nothing, while the real snapshot sits orphaned forever under the old, now-unreachable
ID.

*Reproduction [Executed]:* loaded the pristine app, created an entry, waited 600ms past the
debounce, confirmed via direct IndexedDB inspection that a snapshot existed under
`recovery:lib-fafc322d-…`, reloaded, confirmed a *different* key persisted unread and the
entry list came back empty with "Recovery ready" showing. Reproduced independently with
IndexedDB disabled (localStorage fallback) — same failure. Confirmed the bug does **not**
occur once a library has been saved at least once (a stable concrete `libraryId` from then
on) — a 4th unsaved entry correctly reappeared with "Recovered unsaved changes."

*Affects:* HTML app — `init()` / `resolveRecovery()` / `scheduleRecovery()`.
*Why it matters:* this is precisely the first-session experience of every new adopter, and
precisely the moment they're most likely to close the tab without thinking twice.
*Direction:* mint the library's permanent ID once and embed it into the very first save
immediately, or key the recovery snapshot by a second, session-stable identifier that
survives reloads of the same unsaved file rather than being recomputed on every load.

### HIGH

**H1 — On any browser without the File System Access API, the revision counter and the
"unsaved" indicator both become unreliable.** Firefox, Safari, and any locked-down Chromium
permanently use the download-only save path. Clicking the main Save button twice with an
edit between clicks downloaded two files both stamped `revision:1` with genuinely different
content, because the download path never calls `finishSave()`, so `lastPersisted` never
advances. The "unsaved changes" indicator also never clears.
*Reproduction [Executed]:* diffed two consecutively downloaded files directly.
*Direction:* advance `lastPersisted`/`data.revision`/`dirty` on a successful download the same
way `finishSave()` does for an in-place save.

**H2 — "Uninstall" does not survive the next Word restart.** `AutoExec` unconditionally calls
`SetupClauseLibrary` on every launch; `UninstallClauseLibrary` sets no persisted flag telling
it not to. As shipped, Uninstall only removes the UI until Word is next reopened.
*[Source-reviewed — no persisted disabled-state check exists anywhere in the module.]*
*Direction:* write a marker file (or Settings/registry value) on Uninstall; check it in
`AutoExec`.

**H3 — The visible "Save As" button's no-API fallback never stamps revision/savedAt.** Two
different functions handle "save a copy without a file handle": the main Save button's
fallback correctly calls `makeSavePayload()`; the dedicated `downloadCopy()` — which is what
the "Save As" button itself calls on these browsers — spreads the current in-memory `data`
unchanged, so the download's revision/savedAt are whatever they already were (often `0`/
`null`).
*Reproduction [Executed]:* downloaded file's embedded JSON showed `revision:0, savedAt:null`
after real edits.
*Direction:* route through the same `makeSavePayload()`-based path, or remove the duplicate
function.

### MEDIUM

- **M1** — Cache parsing cost scales with rich-XML volume, not entry count; VBA whole-string
  `Split`/`Replace` has real per-call overhead at large sizes. *[Plausible — browser-side size
  projection executed; VBA timing needs real Word.]*
- **M2** — No pre-flight check for protected/read-only documents; failures surface as raw
  `Err.Description`. *[Plausible — absence confirmed in source; exact text needs real Word.]*
- **M3** — Keyboard-shortcut conflicts fail silently with zero notification.
  *[Source-reviewed.]*
- **M4** — No first-class link between clause variants (§6). *[Design observation.]*
- **M5** — Rich-formatting "stale" detection is exact-string (`richSourceText===clauseText`),
  not substance-aware — a trailing-space edit permanently flips fresh→stale.
  *[Source-reviewed, trivially reproducible.]*

### LOW

- **L1** — Import-conflict resolution doesn't scale past a handful of conflicts (one blocking
  `prompt()` per conflict, no "apply to all"). *[Executed at small scale + source-reviewed.]*
- **L2** — Unbounded growth of `inboxSeenIds`; not a near-term problem, no pruning strategy.
  *[Source-reviewed.]*

### OBSERVATIONS

- **O1** — No injection vector found anywhere in stored-entry rendering. Confirmed by grep
  and two direct adversarial execution probes.
- **O2** — The offline claim holds — zero external requests, zero CDN/font/telemetry
  references anywhere in source.
- **O3** — The inbox/cache split is the right amount of architecture; a heavier design would
  not fix anything this system currently gets wrong. Not a recommendation to change it.

---

## 12. Final assessment

**A. Executive verdict: solid beta.** Not a prototype — the architecture, error handling, and
defensive coding are well past that stage. Not yet a "reliable personal-use tool" because of
one specific, fixable, first-session data-loss bug (C1) a careful audit should not wave
through.

- Browser app alone: beta → near-RC once C1/H1/H3 are fixed.
- Word/VBA integration: well-designed alpha — careful, idiomatic code with the right
  defensive patterns, but categorically unverified end-to-end without a real Word install,
  and H2 is a confirmed real bug.
- Data safety once saved once: high confidence — a full day of adversarial testing (malformed
  files, foreign files, script-injection payloads, import conflicts, scale to 5,000 entries)
  produced zero silent loss or corruption of already-persisted data.

**B. What was actually tested:** see §2 above — every row marked Executed was run in real
Chromium against the exact supplied file; every row marked Source-reviewed was not run.

**D. Word integration, by capability:**

| Capability | Verdict |
|---|---|
| Capturing from Word | Well-designed; revision-safe extraction is a genuine strength **[needs real Word]** |
| Comments | Correctly cannot leak into contract body |
| Tracked changes | Handled thoughtfully on capture; insertion-time interaction untested **[needs real Word]** |
| WordOpenXML round-trip | Correct technique/fallback chain; numbering is the one failure mode that could succeed silently with wrong output **[needs real Word]** |
| Direct Word search | Functional but dated unless the optional workbench is hand-built |
| Rich insertion | Well-guarded with honest warnings before any downgrade |
| Plain-text fallback | Correctly triggered on thrown errors, not on a "successful" wrong-numbering outcome |
| Cache architecture | Sound; watch rich-XML volume at real scale |
| Context menus/shortcuts | Registration idempotent and clean; removal confirmed broken (H2) |
| Corporate M365 deployment | No structural blockers beyond ordinary macro policy; untested in practice |

**E. Data-safety answer:** if you save your library at least once before walking away and
keep saving normally afterward, this audit found no way to silently lose or corrupt that data
across a full day of trying. The one real gap is the window before your very first save,
which the tool does not currently protect.

---

## 13. Recommended next actions

**Fix before serious use:**
- C1 — stabilize the recovery key so it survives a reload before the first save.
- H1/H3 — correct revision/dirty bookkeeping on the no-File-System-Access-API path, since
  that path is every single save for Firefox/Safari users, not an edge case.

**Worth doing soon:**
- H2 — make Uninstall persist across Word restarts.
- M2 — specific error message for protected/read-only documents.
- M3 — one-time notice when a keyboard shortcut can't be claimed.
- M5 — normalize text before comparing for rich-formatting staleness.
- Document the two structural WordOpenXML limitations (cross-range references, numbering)
  somewhere visible before a user captures a numbered defined-terms clause.

**Potential future enhancements:**
- **Lightweight "clause family" grouping** (an additive, optional `familyId` linking
  Preferred/Fallback/Retired variants) — directly addresses the biggest data-model gap found.
  Moderate complexity, low risk. **Recommended.**
- **Batch import-conflict resolution** ("apply to all"). Low complexity.
  **Recommended if multi-user merging is a real use case.**
- **Pre-flight protected-document check.** Low complexity. **Recommended.**
- **A visible warning before the very first save** ("not yet protected against a crash") as a
  stopgap alongside the structural C1 fix. Trivial complexity. **Recommended regardless.**
- **Cache-format version migration path** instead of a hard mismatch rejection. Real but not
  urgent — matters next time the schema changes.
- **Replacing the file bridge with a local server/native host, or a full SQLite backing
  store.** Considered and **not recommended** — would trade a working, well-understood
  failure mode for new install dependencies without fixing anything this audit actually
  found wrong. The JSON-in-HTML model is unusual but is not the source of any finding here.
