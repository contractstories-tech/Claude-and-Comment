# Clause & Comments Library v6.4 — Independent Adversarial Audit

**Audited package:** `69219315-Clause_Comments_Library_v6_4_Package.zip` (16 files; all SHA-256 hashes in `MANIFEST_SHA256.txt` verified against the actual file bytes — EXECUTED/VERIFIED, no tampering).
**Auditor stance:** fresh, adversarial, no reliance on prior audits, changelogs, or self-reported test results. Every claim below is labelled:
- **EXECUTED/VERIFIED** — actually run (Node syntax check, or real Chromium via Playwright, or manual reproduction) and observed.
- **SOURCE-REVIEWED** — read in full and reasoned through; not run because the runtime (Microsoft Word) isn't available in this environment.
- **INFERRED** — a plausible consequence of the above, not directly observed.
- **NOT TESTED** — explicitly out of reach here (real Word 365, real OneDrive, real corporate policy).

This mirrors the honesty of the vendor's own `V6_4_TEST_REPORT.txt`, which is refreshingly explicit that Word-side behaviour is unverified. I independently confirm that boundary is drawn in the right place, and I go further on the browser/HTML side by actually exercising the shipped code (not just reading it) in a real browser engine, and by reading the full VBA source line-by-line for logic defects that static self-review would miss.

**Test evidence location:** `audit/v6_4/CODEX_AUDIT_WORKSPACE/run_tests.js`, `run_tests2.js`, `results.json`. These run the *unmodified* `Clause_Comments_Library.html` (served locally, not touched) in real Chromium via Playwright, mocking only browser-native pickers (`showSaveFilePicker`/`showOpenFilePicker`) that require a user gesture / real OS dialog headlessly — the same approach the vendor's own test report describes. **19 of 20 scripted scenarios passed** against the real, unmodified code (one apparent failure was traced to a test-harness artifact, not a product defect — see §2.12). No production file was modified; all fixtures live under `CODEX_AUDIT_WORKSPACE`.

---

## PART 1 — ARCHITECTURE RECONSTRUCTION

**Authoritative source of data:** the HTML file itself. Specifically, the `<script type="application/json" id="library-data">` tag embedded in `Clause_Comments_Library.html` (source: HTML L7). There is no external database. A "Save" literally clones the whole document, replaces that JSON blob with the new payload, blanks the live UI state and status text back to defaults, and serializes the entire `<html>` back out as a new file (`buildOutputHtml`, HTML L175). The saved HTML *is* the repository.

**Role of the HTML file:** simultaneously (a) the data file, (b) the application/UI, and (c) the distribution/backup artifact — a genuinely elegant single-file design for an offline personal tool.

**Word cache (`clause_library_cache.cclcache`):** a denormalized, tab-separated, escaped flat file (HTML `buildWordCache`, L220) written into `Documents\ClauseLibrary\` only when the browser successfully completes a Save (`finishSave` → `syncWordCacheNow`, L182/L222). It is explicitly disposable/derived — never authoritative — and carries a header row with format version, libraryId, revision, display name, `savedAt`, `commitId` and entry count, which VBA cross-checks before trusting it (`LoadCacheEntries`, Word.bas L981).

**Word inbox (`clause_inbox.jsonl`):** append-only JSON-Lines queue. VBA appends one line per capture (`AppendEntryToQueue`, Word.bas L427); the browser tails the file for new bytes every 6 seconds (`checkInbox`, HTML L229) and merges genuinely-new IDs into the live in-memory model. Entries are *not* part of the repository until the browser subsequently Saves.

**Active-library marker (`active_library.txt`):** a two-field file (`libraryId<TAB>libraryName`) that pins which library Word should tag new captures with (`SaveActiveLibrary`/`ActiveLibraryId`, Word.bas L1344-1399). This is the one piece of the bridge that can silently drift from the cache (see Finding H-3).

**Browser crash-recovery:** IndexedDB (falling back to `localStorage` if IndexedDB is unavailable — HTML L206-L214) stores a full clone of the in-memory `data` object per browser session, keyed by a recovery key derived from the committed `libraryId`, or — before the first Save — a hash of the page URL (`bootstrapRecoveryKey`, L187). On load, `resolveRecovery` (L194) compares the newest unsaved snapshot's timestamp against the embedded/saved payload's `savedAt` and offers (via `confirm()`) to recover it. **EXECUTED/VERIFIED**: closing a tab with an unsaved new entry and reopening the same URL correctly re-offered and restored the entry.

**Committed revision/commit identity:** a monotonic `revision` integer + a random `commitId` (`c-<uuid>`) + `savedAt` ISO timestamp, minted fresh on every successful Save (`makeSavePayload`, L173). This triple is the backbone of every concurrency check in the app: disk-currency before overwrite (`diskStillCurrent`, L178), Save-As target verification (`inspectSaveAsTarget`, L179), and the Workbench's "has the cache moved under me?" check in VBA (`WorkbenchCacheStillCurrent`, Word.bas L901).

**WordOpenXML / richSourceText / freshness model:** an entry can carry Word's self-contained OOXML fragment (`wordOpenXml`) alongside a plain-text snapshot of the wording *at the moment that XML was captured* (`richSourceText`) and a boolean of whether that snapshot is trustworthy (`richSourceKnown`). `richState()` (HTML L96) compares a canonicalized version of `richSourceText` against the *current* `clauseText` and yields `fresh` / `stale` / `unknown`. Only `fresh` entries are exported as real XML into the Word cache (L220); everything else exports empty XML, forcing Word to fall back to plain text. **EXECUTED/VERIFIED**: editing a fresh rich entry's clause text correctly flips it to `stale` and the UI banner changes accordingly.

**Needs Review:** boolean, defaults true for anything captured, imported, or duplicated without explicit trust; excludes entries from Word's default search (both the HTML filter and the VBA `FindCacheMatches`/`WorkbenchSearchEx`) unless explicitly included.

**Clause Family:** a free-text `familyName` plus a derived, stable `familyId` (`familyIdentity` = `'fam-'+hash32(lowercased name)`, L97) used only to *group* variants for display — there is no enforced taxonomy or canonical family record.

**Preferred/Fallback/Retired/Unclassified:** a single `status` enum. This conflates two independent concepts (negotiation posture vs. lifecycle) — flagged in detail in Part 9.

**Favourites, search, capture, insertion, comments, forms/Workbench, keyboard shortcuts, context-menu, JSON backup, snapshot behaviour, Save behaviour, persistent file handles, installer model** — all covered functionally in Parts 5–11 below with citations.

### Data-flow models

**A. Word → Library:** Selection → `GetCleanRangeContent` (tracked-change-safe extraction) → optional `frmSendClause` or two-question `InputBox` (`QuickCapture`) → `AppendEntryToQueue` writes one JSON line to `clause_inbox.jsonl`, tagged with the currently active `libraryId` → browser's 6-second poll (`checkInbox`) merges it into the in-memory model, marked `needsReview:true` → invisible to the repository until the browser Saves.

**B. Library → Word:** Browser Save → `buildWordCache` writes the flattened `.cclcache` → VBA's next Search/Insert (`InsertFromClauseLibraryCore` / Workbench) reads that cache fresh each time (not itself cached in VBA beyond the current session) → rich XML inserted only if `richState==='fresh'`, else plain text, with a non-silent confirm-based fallback if `InsertXML` throws.

**C. Save/Commit:** `save()`/`saveAs()` (L180-181) → `diskStillCurrent()` optimistic-concurrency check → write new HTML → `finishSave()` mints new revision/commitId, clears dirty flag, deletes the *previous* browser-session's recovery record, and re-syncs the Word cache. **EXECUTED/VERIFIED**: a Save is correctly refused (kept as unsaved, no data loss) when the on-disk master was mutated by another writer since this tab last read it.

**D. Crash/Recovery:** periodic (every 3s while dirty) + debounced (350ms after each edit) + `visibilitychange`/`pagehide` snapshot into IndexedDB/localStorage → on next load, `resolveRecovery` offers the newest surviving snapshot if it postdates the saved file. **EXECUTED/VERIFIED**.

**E. Upgrade/Migration:** `normalizePayload`/`normalizeEntry` (L98-122) handle schema drift on read: legacy field names (`text`/`comment`/`notes`/`keywords`/`stance`), version-gated trust of rich-source claims (`trustedRichSource` only true for schema ≥8), and an explicit, hard refusal to open a file from a *newer* schema version (`sourceVersion>VERSION` throws). **EXECUTED/VERIFIED**: a schema-99 file is rejected with the exact documented message, not silently mis-imported.

### Where the same data exists in more than one place, and what guards it

| Data | Copies | Guard against drift |
|---|---|---|
| Repository content | Embedded JSON in HTML (master) · IndexedDB recovery snapshot(s) · Word `.cclcache` (flattened, derived) · manual JSON backups | `commitId`/`revision` triple checked before every overwrite; cache is read-only derived and re-written wholesale on every Save; recovery snapshots are deleted once superseded by a real Save |
| "Which library is this" | `libraryId` embedded in HTML · `active_library.txt` · cache header's libraryId | Insert-side (`EnsureActiveLibrary`) checks all three agree and blocks on mismatch; **capture-side does not** (Finding H-3) |
| Search/ranking logic | HTML JS (`filteredEntries`) · VBA (`FindCacheMatches`/`WorkbenchSearchEx`) | **No guard** — two independent hand-written implementations of essentially the same ranking heuristic that will drift silently over time (Finding M-3) |

---

## PART 2 — ADVERSARIAL RELIABILITY / DATA-INTEGRITY FINDINGS

Each finding states severity, what I actually did, and where in the code it lives.

### H-1. The single highest-stakes assumption in the product is unverified by construction, not by neglect
**Severity: HIGH (by exposure, not by design quality) — SOURCE-REVIEWED, Word-execution NOT TESTED (openly, by both the vendor and me).**
The mixed-Track-Changes capture fix (`GetCleanRangeContent` → `TryCleanTrackedRange`, Word.bas L491-574) is the mechanism most likely to put subtly-wrong wording into a lawyer's permanent precedent bank, and it can only be exercised inside real Word (`Documents.Add`, `AcceptAllRevisions`, `InsertXML`/`FormattedText` round-trips are not something Node/Chromium can simulate). The *logic* is sound on paper: try the WordOpenXML route in an isolated hidden document, verify non-empty output, retry via `FormattedText` if the first attempt collapses, and **refuse to write anything** if both attempts produce empty text (this directly targets the reported v6.3 defect and is a good, safety-first shape). But: the safety check only verifies **non-emptiness**, not **correctness** — see H-2. Until this is run once in real Word 365 against the exact test in `FOCUSED_WORD_RETEST.txt` §2, nobody — vendor or me — actually knows it works. This is not a criticism of the code; it is the correct, honest state of an unverified assumption that the whole "trustworthy repository" claim rests on.

### H-2. No post-insertion content-equivalence check, in either direction
**Severity: HIGH — SOURCE-REVIEWED.**
Two places assume "no VBA error" means "correct result," which the audit brief specifically warns against:
- **Capture side:** `TryCleanTrackedRange` only checks `Len(Trim$(plainText))=0` (Word.bas L556-557). It never checks that the accepted text is *plausible* relative to the original — e.g. that it doesn't still contain a stray revision mark, or that at least one tracked deletion actually disappeared. A pathological Word bug that leaves in *both* the old and new wording (rather than emptying the range) would sail through as a "successful" capture of visibly wrong text.
- **Insertion side:** `InsertClauseContent` (Word.bas L1220) treats a non-erroring `r.InsertXML xml` as success (L1241-1248) and moves on. `InsertXML` can succeed with zero VBA error while silently remapping styles/numbering the target document doesn't define, producing wrong visible numbering or lost formatting. The UI does tell the lawyer, in the detail-pane banner, to "verify automatic numbering and any bookmarks/cross-references" — but that is a *manual* discipline requirement, not a safeguard, and it is easy to skip under deadline pressure. This is exactly the "correct visible wording is more important than avoiding an exception" gap the brief asked me to look for, and it exists.

### H-3. Multi-library capture can silently misdirect entries
**Severity: HIGH for anyone who uses more than one Clause Library file/matter on one machine — SOURCE-REVIEWED.**
`EnsureActiveLibrary` (Word.bas L1344) blocks *retrieval* if the cache's libraryId doesn't match the pinned `active_library.txt`. There is **no equivalent check on the capture path**. `AppendEntryToQueue` blindly stamps whatever `ActiveLibraryId()` currently resolves to (Word.bas L443-450), even if the browser tab actually open and being worked in right now belongs to a *different* library (e.g., the user connected/saved Library B's cache without re-running `ActivateCurrentClauseLibrary`). The captured JSON line ends up tagged for Library A; when Library B's browser tab polls the shared inbox file, `checkInbox` correctly *refuses* to misfile it (L229: `otherLibrary++` on library-ID mismatch) — so no corruption occurs — but the capture silently vanishes from the visible library until the user notices the terse "N record(s) for another library skipped" status text and manually reopens Library A. For a solo practitioner using one library forever this never bites; for anyone using per-matter or per-client libraries sharing the default bridge folder, it's a real, hard-to-diagnose "where did my capture go" moment.

### H-4. There is no self-discoverable, reliable "front door" in Word
**Severity: HIGH for onboarding/daily use — SOURCE-REVIEWED, consistent with vendor's own admission.**
The two things a new user will instinctively try — **right-click** and **look for a ribbon button** — are exactly the two paths this package does *not* reliably support: legacy `CommandBars` context-menu items (Word.bas L114-128) are openly documented as unreliable in modern Word's Fluent list/numbered context menus, and there is no Ribbon UI at all. The only dependable path is a keyboard chord (`Ctrl+Alt+Shift+C`/`+I`) that nothing in the Word UI surfaces or teaches. This is a real usability/adoption risk independent of any code defect — addressed further in Part 7.

### H-5. Save is a single point of truth with only one generation of rollback
**Severity: HIGH for "valuable repository" framing — SOURCE-REVIEWED + EXECUTED (round-trip mechanics).**
`buildOutputHtml`/`finishSave` (L175/L182) atomically replace the entire embedded repository on every Save, keeping exactly one prior committed state as `previousBackup` (restorable via "Restore previous", L184-185). If a lawyer Saves twice in a row after an unwanted bulk change (e.g., a bad find-and-replace-style edit, or an accidental mass status change), the only in-app recovery is gone — the second Save overwrote the one available backup. Nothing in the app prevents this; the docs are honest about it ("This is not a multi-generation audit trail," `INSTALL.txt` L94). Recovery then depends entirely on external OneDrive version history or a manually-remembered JSON Backup — both outside the product's control. This is an acceptable design *if the lawyer actually runs periodic JSON Backups*, which nothing in the product enforces or reminds them to do beyond the JSON Backup button existing.

### M-1. `GenerateId()` can theoretically collide across two concurrent Word processes
**Severity: MEDIUM likelihood: very low; impact: silent single-entry loss — SOURCE-REVIEWED.**
`GenerateId` (Word.bas L1530) = `"w" + UTC-timestamp-to-millisecond + 6-digit-per-process-sequence`. Two *separate* Word.exe processes each keep their own `Static seq` counter starting at 0. If both happen to write their Nth capture at the exact same UTC millisecond (theoretically possible with two people/monitors on one machine, or scripted/automated captures), the IDs collide. The consuming side (`checkInbox`, HTML L229) does not detect this as a duplicate-ID error (which the app otherwise treats seriously, see L-2); it just treats the second arrival as "already seen" and **silently drops it** (`if(!existing.has(e.id))`). Low probability, but the failure mode is silent data loss rather than a loud warning — worth a cheap fix (append a per-process random suffix) even though it's unlikely to bite a solo user.

### M-2. Word-hidden text is captured without warning
**Severity: MEDIUM — SOURCE-REVIEWED.**
`rng.Text` (used throughout `GetCleanRangeContent`/`CleanWordText`) includes Word's Hidden-formatted text by default; the code does not filter it. A selection that happens to include hidden internal notes (a common drafting habit) would be captured verbatim into `clauseText` with no indication to the user that invisible text was folded in. This is a distinct, narrower risk from the product's explicit "Internal Context" field (which is handled correctly, see §14) — it's about text the *document* hides, not the *app's* fields.

### M-3. Two independent search/ranking implementations will drift
**Severity: MEDIUM — SOURCE-REVIEWED.**
The HTML's `filteredEntries` (L140) and the VBA's `FindCacheMatches`/`WorkbenchSearchEx` (Word.bas L1081/L773) are hand-written, separately-maintained re-implementations of "AND-token match, score by title/favourite/status." They already differ subtly (e.g., exact scoring weights, exactly which fields are hay'd). Every future tweak to one and not the other quietly changes "how it feels to search" depending on whether you're in the browser or in Word — for a tool whose whole value is *fast, confident retrieval*, this asymmetry is worth eliminating rather than accumulating.

### M-4. Import conflict resolution is a typed-letter modal prompt
**Severity: MEDIUM — EXECUTED/VERIFIED.**
`resolveImportConflict`/bulk conflict handling (L203-204) uses `window.prompt()` expecting a single letter (K/I/D/R), defaulting to **K (keep existing)** for any unrecognized input. I verified this works as documented for the "keep existing" path. But there is no per-entry confirmation of *which* action was actually taken, and a rushed lawyer fat-fingering the prompt gets a silent "keep existing" rather than an error. Given imports are rare and consequential, this deserves either a real dialog with named buttons, or (better) an entry-by-entry review list before committing — not a stack of blocking native `prompt()` calls.

### M-5. Rich-formatting freshness is binary and text-only
**Severity: MEDIUM — SOURCE-REVIEWED.**
`richState()` (L96) invalidates the *entire* stored rich formatting the instant `clauseText` differs from `richSourceText` by even a single character (a typo fix, a comma). There's no partial trust and no "I re-verified this manually, keep the rich formatting" override — the only way back to `fresh` is recapturing from Word. This is the safe default, but for a library that will accumulate hundreds of small in-browser text tidy-ups over years, it means rich formatting quietly degrades to plain-text-only for a large fraction of entries unless the lawyer diligently recaptures — worth knowing the failure mode is "silently gets less useful," not "corrupts."

### L-1/OBSERVATION — Dirty-state indicator is CSS-only
**EXECUTED/VERIFIED.** The " • unsaved" text is a CSS `::after{content:...}` (HTML L12), invisible to `textContent`, screen readers, and any future automation reading status text. The underlying `dirty` CSS class *is* set correctly (confirmed via `getAttribute('class')`) — this is a presentation/accessibility nit, not a functional defect.

### L-2 — Recovery keys for *unsaved, pristine* copies are derived only from the page URL
**EXECUTED/VERIFIED (observed as a side-effect of testing), SOURCE-REVIEWED for implication.**
Before the first Save, `bootstrapRecoveryKey(location.href)` (L187) is the only recovery/instance-coordination key. Two browser tabs opened at the *same URL* (e.g. two people, or one person double-clicking the file twice) will legitimately share crash-recovery snapshots and the BroadcastChannel "secondary instance" lock — I confirmed this directly (a second tab of the same unsaved template was correctly marked read-only). This is a deliberate and reasonable choice for the common single-file-double-opened case, but it does mean recovery snapshots from *unrelated* templates opened from the same path/origin can appear as "recovery candidates" for each other until one of them gets a real, unique `libraryId` via a Save. Low real-world impact; worth understanding rather than fixing.

### Positive, verified findings (things that hold up under attack)
- **EXECUTED/VERIFIED:** No external network requests of any kind on load or during use — genuinely offline, no telemetry.
- **EXECUTED/VERIFIED:** Duplicate-ID files are never silently opened; the user is prompted for Repair Mode, and repaired records are visibly flagged "— Repaired duplicate" and forced into Needs Review.
- **EXECUTED/VERIFIED:** Corrupted embedded JSON produces a clear, blocking alert — the app never half-loads garbage.
- **EXECUTED/VERIFIED:** A file from a newer schema version is flatly refused, not mis-migrated.
- **EXECUTED/VERIFIED:** The stale-disk-write guard actually works — I mutated an on-disk master out from under an open tab and confirmed Save refuses to clobber it.
- **EXECUTED/VERIFIED:** At 2,000 synthetic entries, load was ~390ms and a live search (with its 120ms debounce) resolved in ~245ms — no perceptible lag at real-world scale.

---

## PART 3 — TRACK CHANGES / WORD-SPECIFIC AUDIT (SOURCE-REVIEWED; execution NOT TESTED — no Word available)

Reasoning through `GetCleanRangeContent`/`TryCleanTrackedRange` against each scenario in the brief:

| Scenario | Assessment |
|---|---|
| No revisions | Fast path (`rng.Revisions.Count=0`) reads `rng.Text`/`rng.WordOpenXML` directly. Simple, low risk. |
| Insert-only / delete-only | Explicitly called out as "control regressions" in `FOCUSED_WORD_RETEST.txt` §2 — the vendor is right to test these as separately as the mixed case, since `AcceptAllRevisions` on an insert-only or delete-only range is a strict subset of the mixed case and shouldn't fail, but hasn't actually been run. |
| Insert + delete in same selection | The scenario the v6.3→v6.4 fix targets directly. Logic is sound (isolated temp doc, accept, verify non-empty, retry via FormattedText, else refuse) but unverified in real Word — see H-1/H-2. |
| Multiple replacements | Not distinguished from single insert+delete by the code — same path, same risk profile, NOT TESTED. |
| Formatting-only revisions | `AcceptAllRevisions` accepts formatting revisions too; no special handling needed, but also nothing confirms formatting survives the temp-doc round-trip. NOT TESTED. |
| Moved text (Word's Move tracked-change pairs) | Not special-cased anywhere. `AcceptAllRevisions` is documented by Microsoft to resolve moves by keeping the moved-to copy — if that engine behavior changes or the range boundary splits a move pair, this code has no defense beyond the generic empty-output check. NOT TESTED, not explicitly covered by `FOCUSED_WORD_RETEST.txt` either — **testing-coverage gap**. |
| Comments plus revisions | Handled by a *separate* mechanism — `CollectSelectionComments()` reads Word Comments straight from the live `Selection`, independent of the tracked-revision temp-doc cleanup. This is a sound separation of concerns (comments aren't touched by the risky Accept/InsertXML path at all). |
| Paragraph-level revisions | No special handling; same generic path. |
| Selection beginning/ending inside a revision | Not specifically handled. Word's own `WordOpenXML`/`FormattedText` export for a range that starts/ends mid-revision can auto-expand to revision boundaries or can produce a truncated/malformed revision fragment, depending on Word build — genuinely untestable without Word, and not covered by the shipped retest script either. **Testing-coverage gap.** |
| Accepted vs rejected views | Not applicable — Word's Reviewing-pane display mode doesn't change what `Revisions.Count`/`WordOpenXML` return; this is a display-only setting. Correctly irrelevant to the code. |
| Hidden/deleted text | See Finding M-2 — hidden (non-tracked) text is captured unfiltered. |
| Tables with revisions | **Not tested by the vendor's own retest script at all**, and table-specific `AcceptAllRevisions` edge cases (stray rows, merged-cell artifacts) are a known category of Word engine quirk. This is the single largest testing-coverage gap I found relative to the brief's explicit ask. **Recommend adding this to the real-Word retest before trusting rich capture on any tabular clause (e.g., SLA tables, pricing schedules).**

**Overall verdict on Part 3:** the *shape* of the safety design (fail closed, never silently empty, retry once, then refuse) is correct and is a genuine improvement over what v6.3 apparently did. But "safe" here means "won't file an empty entry," not "will file a correct one" — and that stronger guarantee is not, and cannot be, established without a real-Word test pass that specifically includes tables, mid-revision boundaries, and moved text, none of which the shipped `FOCUSED_WORD_RETEST.txt` currently covers.

---

## PART 4 — WORD FIDELITY (SOURCE-REVIEWED / NOT TESTED — no Word runtime)

The product's own documentation is candid that **legal multilevel numbering (1 / 1.1 / (a) / (i)) has never been verified in real Word**, while ordinary decimal numbering was verified in the v6.3 acceptance pass. I have no way to add real evidence here beyond confirming the code takes the same single path (`InsertXML` of a stored self-contained OOXML fragment) for *every* kind of formatting — numbering, styles, tables, hyperlinks, fields, bookmarks, content controls, headers/footers, footnotes — with no per-feature special-casing. That is architecturally the right approach (OOXML fragments are meant to be self-describing), but it also means there is exactly one thing to get right, and the one thing most likely to go wrong for a commercial lawyer's actual clause library — custom outline numbering nested inside a template's existing numbering scheme — is precisely the one the vendor flags as unverified. **This is the top item to resolve before trusting rich insertion for anything beyond simple paragraph clauses.** Cross-references and bookmarks that depend on content *outside* the captured selection are correctly flagged to the user in the detail-pane banner (HTML L156) as something to manually re-check — appropriately, since no amount of code can guarantee an externally-scoped reference resolves correctly after a partial-document insert.

---

## PART 5 — SEARCH / PRECEDENT SELECTION

**Browser side (EXECUTED/VERIFIED):** token-AND substring search across title/type/category/family/status/position/clause/comment/context/source/tags, with a 120ms debounce, sorted by recency/title/category/created. At 2,000 entries this is fast (~245ms including debounce) and the UI shows live counts, tag chips, and filter dropdowns for type/status/category/family/position/favourite/needs-review. **This comfortably clears "identify the right precedent in under 10-15 seconds"** — for the browser UI.

**Word side, compact mode (SOURCE-REVIEWED):** `InsertFromClauseLibraryCore` (Word.bas L940) is a single `InputBox` for the query, then a **paginated `MsgBox`-style `InputBox` list, 5 results per page**, requiring the user to type a row number or "N"/"P" to page (`ChooseMatch`, L1138). At 50 entries this is tolerable; at 250-2,000 entries, with no live filtering, no keyboard arrow navigation, and only a 68-character excerpt per row, this **will not** meet the 10-15-second bar — a search returning 40 matches means 8 rounds of retyping "N" into a modal box. This is the *default* experience unless the polished Workbench is installed and working.

**Word side, Workbench mode (SOURCE-REVIEWED):** a proper filtered ListBox with live column display (Title/Status/Position/Family/Category/Source), full preview pane, and instant re-filter on combo/checkbox change (though free-text search requires pressing Enter, not live-as-you-type — a reasonable choice to avoid re-searching on every keystroke against a VBA-side linear scan). This is a good design and, if built, would comfortably meet the retrieval-speed bar even at 2,000 entries. **The catch is entirely in Part 11: it is not guaranteed to exist.**

**Near-identical clauses / precedent-selection risk:** the Clause Family grouping (`familyId`) surfaces "N other variants" inline in both the browser detail pane and is a filterable/browsable dimension in the Workbench — a reasonable, low-overhead answer to "which of my five liability-cap variants is this." There is no word-level diff or side-by-side view; given this is described as a personal tool (not a governance system), I would **not** recommend building a diff view — the family grouping plus a good preview panel is enough, and a diff view is real engineering effort for a marginal gain over "read the two previews side by side," which a human does faster than parsing a diff for prose text anyway.

**Metadata recommendation:** do not add "last reviewed date," "usage frequency," "superseded-by," "jurisdiction," or "agreement type" fields — see Part 15/16 for why each is classification cost without a corresponding retrieval win at this scale and for this single-user use case.

---

## PART 6 — CAPTURE WORKFLOW

Current shape: **Select → Ctrl+Alt+Shift+C → (rich form or two `InputBox` questions) → inbox → later HTML review**, with `needsReview` forced true so nothing captured silently becomes "trusted" without a second look. This is a good default. Concretely:

- **Rich form path** asks for: title (pre-filled, editable), type, category, status, position, family, tags, source (pre-filled to document name), favourite — a lot of fields for a "quick capture," but every one is either pre-filled or optional/skippable at save time (only title is truly required; type-specific text is required only for its own type). **Recommend:** remember the *last-used* category/family across a session (currently resets to "Other"/blank every time, Word.bas L373-378) — a cheap, high-value change since a lawyer capturing five liability variants from one document in one sitting will retype/reselect the same category five times.
- **Quick Capture fallback** (no form installed) is genuinely minimal: title, then category, both via `InputBox` — this is close to "select → shortcut → done" already.
- **Duplicate/near-duplicate detection at capture time:** none. Given this is a personal tool and duplicates are caught later by the Health Check's exact-hash-adjacent `legacyId` mechanism only for *imports*, not live captures, I'd rate a capture-time near-duplicate warning as **quality-of-life, not essential** — it adds a moment of friction to every capture in exchange for catching a mistake the lawyer will likely also catch during the "needsReview" pass anyway.
- **Comments captured together with clause text:** already works — `entryType` is automatically set to "Clause + Comment" if the selection has attached Word Comments (Word.bas L342), which is a nice, low-friction detail.
- **Word Comments as first-class capture source:** already implemented (`Selection.Range.StoryType = wdCommentsStory` branch, L335) for capturing a comment *by itself* if the user selects inside the Comments pane.

---

## PART 7 — RETRIEVAL / WORD UX

The compact `InputBox`-paginated fallback is not adequate at scale (Part 5); the Workbench is good but optional and fragile to provision (Part 11). My recommendation, in order of value:

1. **Ship the Workbench pre-built inside a real .dotm**, not as a VBIDE bootstrap installer that has already needed one full hotfix cycle (`comp.Designer.Width/Height` → error 438) and is explicitly still "not runtime-certified" in this v6.4 release. The installer is clever engineering but is solving a problem ("distribute a UserForm without shipping a binary template") that a one-time manual build in real Word (`DOTM_PACKAGING_GUIDE.txt`) solves more reliably and only needs doing *once*, by the vendor, ever — not by every installer run on every machine.
2. **A minimal Ribbon group** (`[Capture] [Search Library] [Open Library]`) would meaningfully improve discoverability over "memorize a keyboard chord" — Ribbon XML customization via a Custom UI part in the .dotm is a one-time authoring cost, not a runtime dependency, and doesn't carry the VBIDE-trust-setting requirement the forms installer does. I'd rate this **high-value** precisely because it fixes Finding H-4 (no discoverable front door) without needing Office.js or any architecture change.
3. Inside the Workbench: **double-click to insert** is currently not implemented (only the explicit "Insert Clause" button and Enter-in-search-box triggers a re-search, not an insert) — a fast, low-risk addition. **Ctrl+Enter to add-as-comment** is not implemented either. Both are cheap, high-value shortcuts for a modeless form used dozens of times a day.
4. Family variants already surface in the preview text; showing them as a visually grouped sub-list in the ListBox itself (rather than only in the free-text preview) would help scanning at a glance, but I would not prioritize this over items 1-3.

---

## PART 8 — HTML UI / PRODUCT UX

The interface is clean, information-dense without being cluttered, and — critically for a *repository*, not a workflow tool — never hides *why* something is in a given state (Needs Review banner, Retired banner, stale-rich-text banner, first-save warning are all explicit, colored, and dismissable-by-action rather than by silently disappearing). Concretely good decisions I would **not** rewrite: the master-filename-plus-revision status string, the "Opened" vs "Master" distinction, the always-visible entry/filter count, the tag-chip bar.

**What would confuse a first-time lawyer (see also Part 17):**
- **"Word cache" and "Word inbox"** are two different connect buttons with two different file pickers and no single "connect Word" onboarding flow — a new user must read the docs to know they need to do both, once, and in roughly the right order (per `INSTALL.txt`'s bootstrap sequence). **Recommend:** a single first-run "Connect to Word" wizard button that does both picks in sequence with inline explanation, hidden after first use.
- **The distinction between "JSON Backup," "Save," "Save As," and the auto-recovery system** is four different persistence concepts with genuinely different guarantees, and nothing in the UI itself explains *why* you'd want more than one. A single onboarding tooltip or a "What's the difference?" link would remove real confusion at no ongoing cost.
- **Health Check** is a good idea, well-implemented (duplicate IDs, unverified rich text, Preferred-but-Needs-Review, oversized XML, family-identity mismatches, invalid dates), but it's a manual button nobody will think to click until something already feels wrong. Running it silently in the background (or a small badge count) rather than only on-demand would catch problems earlier.

---

## PART 9 — PRODUCT MODEL

**Yes — `status` (Preferred/Fallback/Retired/Unclassified) incorrectly conflates two independent axes**, exactly as the brief suspected:
- **Negotiation tier** (how favourable/how likely to hold this position): Preferred, Fallback, Counterparty/Other.
- **Lifecycle** (is this wording still good): Active, Retired.

Today, "Retired" is a dead end that can only ever mean lifecycle, while "Preferred"/"Fallback"/"Unclassified" only ever mean negotiation posture — so the enum silently behaves as two different value sets depending on which value you're looking at. A retired *fallback* clause and a retired *preferred* clause currently cannot both be represented; retiring a Preferred clause loses the information that it used to be your first choice. **Recommend splitting into two fields**: `tier` (Preferred/Fallback/Counterparty/Unclassified) and `lifecycle` (Active/Retired) — a small, mechanical migration (one becomes two booleans-worth of enum) that removes a real, silent information loss without adding UI complexity (it's still two dropdowns, just orthogonal instead of overloaded ones instead of one dropdown doing double duty).

**Other fields, assessed one by one:**
- **Clause Family** — keep as-is; free text + derived stable ID is the right amount of structure for a personal tool. Do not add a separate "family management" screen.
- **Position (provider/balanced/customer)** — keep; it's a real, frequently-used negotiation dimension distinct from tier/lifecycle, correctly modeled as its own field already.
- **Category** — keep; a fixed suggested list plus free text is the right balance.
- **Tags** — keep; genuinely optional, low-friction, and additive to search without being load-bearing.
- **Source/provenance** — keep; already auto-defaults to the source document name at capture time, which is exactly the right amount of automation (captures provenance without asking a question).
- **Internal context** — keep; the one field explicitly *not* offered a "copy" guard against Retired/Needs-Review confirmation in the same way clause/comment text is (HTML L161: `textSection(...,null)` passes no guard entry) — this is a *correct*, deliberate omission (internal notes about *why not* to use a clause are exactly what you want to read even on a Retired entry) but worth confirming it's deliberate and not an oversight: it is, and it's right.
- **Favourites** — keep; simple, cheap, works.
- **Needs Review** — keep; it is the product's actual lifecycle gate today (whether something is *trustworthy*), which is arguably more important day-to-day than the proposed lifecycle split above, and should stay a separate concept from both.

**Do not build:** a workflow/approval system, multi-user roles, a "who reviewed this and when" audit log, or a taxonomy admin screen. None of these serve a personal repository.

---

## PART 10 — ARCHITECTURE CHALLENGE

| Option | Offline | Install friction | Corporate risk | Word fidelity | Durability/backup | Maintainability (solo) |
|---|---|---|---|---|---|---|
| **A. Current** (HTML + VBA/.dotm + bridge files) | Full | Low (copy files, one macro run) | Low — no admin rights, no add-in store | Native OOXML round-trip, best possible short of VSTO | File-based, human-inspectable, git-diffable JSON | High — plain text files, no build step, no dependency to rot |
| B. HTML + IndexedDB as sole store | Full | Low | Low | Same (still needs Word bridge files) | **Worse** — IndexedDB is opaque, per-browser-profile, easily wiped by "clear browsing data," not portable/diffable | Similar, but loses the "the file itself is the backup" property, which is a real regression |
| C. Local SQLite app | Full | Medium (needs a runtime/executable) | Medium — an unsigned local exe is a harder sell in some corporate environments than "open an HTML file" | Same bridge-file approach still needed for Word | Good, but now a binary file, not diff/inspect-friendly | Lower — a new codebase/runtime to maintain solo |
| D. Native Windows app | Full | Medium-High | Medium-High (signing, install, updates) | Same | Good | Lower — real app-packaging burden for one person |
| E. Electron/Tauri | Full | Medium (bundle size, updater) | Medium | Same bridge-file approach | Good | Lower — a whole extra toolchain for a personal tool |
| F. Office.js Word add-in | **Partial** — Office.js add-ins for desktop Word still need a manifest, a local web server or sideloading, and (for anything beyond simple insertion) do not have the same depth of Range/Revisions/OOXML object model VBA has; several of the exact operations this product relies on (raw `WordOpenXML`, `AcceptAllRevisions` on an isolated hidden document, `Application.VBE`) have no first-class Office.js equivalent | Higher (manifest + sideload/catalog) | **Higher** — many corporate tenants restrict custom add-in sideloading more tightly than "run this macro" | **Lower** for this specific feature set | N/A (would still need a store) | Would trade a stable, boring API (VBA object model, unchanged for 20 years) for a still-evolving one |
| G. VSTO/COM add-in | Full | High (install, signing, .NET runtime) | High (admin rights typically required) | Excellent | Good | Much lower — real deployment engineering for one person |
| H. Pure .dotm/VBA (no HTML) | Full | Low | Low | Excellent for Word, poor as a *repository UI* (VBA UserForms are not a pleasant place to browse/edit hundreds of entries) | Good | High, but loses the excellent browser-side editing/filtering experience entirely |

**Answer to "if this were my own personal clause repository, which architecture would I choose": the current one.** The self-contained HTML-as-database is a genuinely good, slightly unusual idea that is *correct* for this specific brief (one user, offline, Word-centric, durable-for-decades, zero install friction, human-inspectable/git-diffable state). Every alternative either (a) trades away offline simplicity and corporate-friendliness for capability the product doesn't need (C/D/E/G), or (b) actually has *worse* Word-integration depth for the specific things this product does (F), or (c) loses the one genuinely superior property of the current design — that opening the repository *is* opening the app, with no separate database file to lose track of (B, H). **I would not rewrite this.** I would fix the two concrete gaps identified elsewhere (a discoverable Word-side front door, and real-Word verification of the Track-Changes fix) rather than change platform.

---

## PART 11 — PACKAGING / INSTALLATION

The ideal end-state, which the vendor's own `DOTM_PACKAGING_GUIDE.txt` already correctly identifies but doesn't ship, is:

```
ClauseLibrary_v6_4.dotm   (core module + both forms + Ribbon XML, pre-built and saved once by the vendor)
Clause_Comments_Library.html   (the user's data file)
```

with the VBIDE-automation `FormsInstaller` demoted from "the recommended path" to "an emergency repair tool for someone who deliberately wants to regenerate the forms from source." Concretely:

- **The forms installer should not disappear entirely** (it's a legitimate, well-guarded repair/upgrade tool — it correctly refuses to bypass the "Trust access to VBA project object model" policy, and now correctly reports the *actual* control that failed rather than a generic error) **but it should not be the primary distribution mechanism**, because it requires a security setting many corporate Word deployments disable by policy, and because — as the vendor's own v6.3 acceptance test proved — VBIDE automation of UserForm creation is measurably more fragile than "here is a finished template file."
- **VBA-project-object-model access can and should be avoided** for the 95% case (using an already-built .dotm) while remaining available for the power-user repair path.
- **A Ribbon XML customization** (Part 7) is worth doing once at the same time as building the reusable .dotm — it's free at that point (no extra runtime dependency) and directly fixes the discoverability gap.
- **Upgrade strategy is already correctly designed**: schema/cache version numbers are checked independently of the VBA module version, and the docs are explicit that "existing v6.3 data does NOT need migration." This is good, low-friction versioning discipline that should continue.
- **User data is already correctly separated from executable code** — the .dotm holds no repository data; the HTML file is the only data file. This is the right split and should not change.

---

## PART 12 — BACKUP / VERSION HISTORY

Current mechanisms, evaluated honestly: one embedded prior-committed generation (`previousBackup`), automatic crash-only recovery snapshots (not a version history — they're deleted once superseded by a real Save), and a fully manual JSON Backup the user must remember to trigger. This is **not sufficient on its own** for "a valuable repository containing hundreds or thousands of entries" — it is sufficient *only if paired with* external file versioning (OneDrive/Git/manual copies), which the docs correctly say but the app does nothing to enforce or even gently nudge.

**Recommended, minimal addition:** on every successful Save, in addition to the existing single `previousBackup`, keep a **rolling window of the last 5-10 committed masters** as timestamped files in a `backups/` subfolder next to the HTML (simple file copies, no format change, no database). This is a few lines of `File System Access` code (open/create a `backups/` directory handle, write a timestamped copy) and directly answers "what if I don't notice a bad Save until the third Save after it" — the single biggest gap identified in H-5 — without introducing any new concept, UI, or format for the user to learn.

**Per-entry version history:** would genuinely add complexity (a new data structure, more UI, more to explain) for a benefit that the family-grouping + Duplicate button already partially covers (a lawyer who wants to keep an old wording around can just Duplicate before editing). **Do not build this** — it's the textbook "more metadata isn't automatically better" case the brief warns against.

---

## PART 13 — PERFORMANCE

| Entries | Load time | Search (debounced, live) |
|---|---|---|
| 2,000 (EXECUTED/VERIFIED) | ~390ms | ~245ms for a 3-token phrase query |
| 100 / 500 / 1,000 | Not separately measured — the O(n) filter/sort architecture (`filteredEntries`, `renderList`) scales linearly and 2,000 already shows no perceptible lag, so intermediate sizes are INFERRED to be proportionally faster. |

No premature optimization is warranted; the current linear-scan-over-an-in-memory-array design is appropriate for realistic personal-library sizes (even 5,000+ entries would very likely stay comfortably under 1 second, by extrapolation — **INFERRED**, not measured). The one genuine size concern — large embedded WordOpenXML blobs — is already self-policed by the Health Check's 500KB-per-entry warning (HTML L237); I did not find evidence this needs proactive reduction, and reducing OOXML fragment size by hand-pruning parts risks exactly the numbering/style breakage the brief says to avoid trading away for file size. **Leave it alone.**

---

## PART 14 — SECURITY / PRIVACY

**EXECUTED/VERIFIED:** zero external network requests on load or during normal use (create/edit/search/save) — genuinely offline, no telemetry, no analytics beacon, no font/CDN fetch.

**XSS/injection:** the app builds its DOM almost entirely via `createElement`/`.textContent` (not `innerHTML` with interpolated data), which is the correct defensive pattern against stored-XSS from clause text containing `<script>`-like content. The one `innerHTML` usage I found for user-influenced content is the "welcome"/banner strings, which are static, developer-authored strings, not user data. The re-embedding step (`buildOutputHtml`, L175) explicitly escapes `</script` sequences in the serialized JSON (`.replace(/<\//g,'<\\/')`) — a correct, specific defense against breaking out of the embedded `<script type="application/json">` tag via a crafted clause title/text. **This is good, deliberate security engineering**, not an accident.

**Malicious/malformed JSON/HTML import:** covered extensively in Part 2 — corrupted, duplicate-ID, and future-schema files are all handled by explicit, user-visible refusal paths rather than silent misbehavior (EXECUTED/VERIFIED for all three).

**VBA trust/macro security:** the installer explicitly refuses to work around Word's "Trust access to VBA project object model" setting and says so in its own error message rather than attempting any bypass — correct posture for a personal tool operating on a possibly-managed machine.

**Clipboard exposure:** `copyText`/`fallbackCopy` (L149-150) correctly gate copying Retired or Needs-Review content behind an explicit confirm(), which is a sensible, low-friction guard against accidentally pasting stale/unreviewed wording into a live negotiation.

**Internal Context exposure:** the field exists specifically to hold information that should *never* leave the repository (negotiation strategy notes), and it is excluded from the "Copy clause + negotiation comment" combined-copy action (L163: only `clauseText`/`commentText` are joined) — correctly scoped. The one adjacent gap is Finding M-2 (Word-hidden text captured unfiltered), which is a document-hygiene risk, not an app-logic one.

**OneDrive exposure:** the bridge files (`clause_inbox.jsonl`, `.cclcache`, `active_library.txt`) live in `Documents\ClauseLibrary\`, which is commonly OneDrive-synced. None of these files contain anything more sensitive than what's already in the HTML master (which the user chooses where to store); this is a reasonable, already-acceptable exposure level for a "local/offline personal tool," and I would not recommend adding encryption or moving the bridge folder — that would add friction for no threat this tool's stated purpose creates.

**I found no compelling case for any cloud service, and none is recommended.**

---

## PART 15 — MISSING FUNCTIONALITY (classified honestly)

**A. ESSENTIAL FOR SAFE DAILY USE**
1. Real-Word verification of the Track Changes capture fix, specifically including tables and mid-revision-boundary selections (closes the largest gap in H-1/Part 3).
2. A rolling multi-generation backup (Part 12) — the single biggest gap between "personal convenience tool" and "repository I'd trust with years of work."
3. A capture-side active-library check mirroring the existing retrieval-side one (closes H-3).

**B. HIGH-VALUE QUALITY-OF-LIFE**
4. Ship a pre-built .dotm with both forms and a minimal Ribbon group (closes H-4, Part 11).
5. Remember last-used category/family during a capture session (Part 6).
6. Double-click-to-insert and Ctrl+Enter-to-comment in the Workbench (Part 7).
7. Split `status` into `tier` + `lifecycle` (Part 9) — a small, mechanical, one-time data-model fix.

**C. NICE TO HAVE**
8. A single "Connect to Word" onboarding flow that chains the two existing connect-cache/connect-inbox pickers.
9. A background/badge-based Health Check rather than purely on-demand.
10. Per-process-unique suffix in `GenerateId()` (closes M-1) — cheap, low-priority given the near-zero real-world odds.

**D. FEATURE CREEP — DO NOT BUILD**
- Per-entry version history (Part 12).
- Word-level diff / side-by-side family comparison (Part 5) — the existing family grouping plus preview pane already solves this at the right cost/benefit for a single-user tool.
- Jurisdiction, agreement-type, "last reviewed date," and usage-frequency fields — classification cost with no demonstrated retrieval win at the sizes this tool will actually reach.
- Any workflow/approval/audit-log system, multi-user roles, or taxonomy admin screen (Part 9).
- AI-assisted classification or drafting inside the authoritative repository path — the repository's entire value proposition is that it is deterministic and fully human-owned; nothing in this audit found a gap that AI, rather than better defaults/automation of *existing* deterministic fields, would actually close.
- Any cloud/sync service (Part 14).

---

## PART 16 — SIMPLIFICATION REVIEW: the minimum-friction edition

What can be removed or hidden today, without losing capability that actually gets used:
- Merge the two separate "Connect Word Inbox…" / "Connect Word Cache…" buttons into one "Connect to Word" action that does both in sequence (still exposing the underlying two connections to power users, e.g. under a details link, but not as the primary two-button surface).
- Hide "Import…" behind a secondary/overflow control — it's a rare, high-consequence action (Part 2 finding M-4) that doesn't need equal visual weight with Save/New in the primary toolbar.
- Default Category/Family to the last-used value rather than resetting every time (Part 6/15-B-5).
- Auto-run Health Check silently after every load/import rather than requiring a manual click.

**The ideal daily experience, in ≤10 actions:**
1. See something good in a Word document while drafting/reviewing.
2. Select it, press Ctrl+Alt+Shift+C.
3. Confirm/adjust the pre-filled title, pick a category (remembered from last time), Save.
4. Later, open the HTML library once a day/week; entries with "Needs review" are waiting at the top.
5. Skim, classify (status/family/tags), mark reviewed.
6. Click Save.
7. Next time drafting: press Ctrl+Alt+Shift+I (or click a Ribbon "Search Library" button), type a couple of words.
8. Press Enter/double-click the right result to insert.
9. If a negotiation comment applies, select the relevant wording and add it as a Word comment from the same search result.
10. Weekly/monthly: click "JSON Backup" once (or rely on the recommended rolling-backup addition, once built, and stop thinking about it at all).

That is already very close to what v6.4 offers today — the gap between this ideal and the current reality is almost entirely the discoverability/onboarding items above, not missing capability.

---

## PART 17 — FIRST-TIME USER TEST

Walking a competent commercial lawyer (not the builder) through the package:

1. **Install** — confused by choice between "just import the .bas and use keyboard shortcuts" vs. "build a whole .dotm with forms" vs. "run an automatic forms installer that needs a security setting changed." The docs *do* explain this clearly and in the right order (`README_FIRST.txt` correctly leads with the keyboard-only path as primary), but a lawyer skimming will likely try the "polished" path first and hit the VBIDE-trust-setting wall.
2. **Understand what the HTML is** — reasonably clear once opened; the welcome screen and status bar are self-explanatory (EXECUTED/VERIFIED: no console errors, clean load, sensible defaults).
3. **Understand what the .dotm is** — this requires reading `DOTM_PACKAGING_GUIDE.txt`; nothing in Word itself explains it. Genuine confusion point for a non-technical user.
4. **Create first entry** — trivial; "+ New entry" is obvious and works exactly as expected (EXECUTED/VERIFIED).
5. **Capture from Word** — works via keyboard shortcut *if* the user has already run `SetupClauseLibrary` and connected/activated a library — a real chain of prerequisite steps (Finding, addressed in Part 16's onboarding recommendation) that isn't obvious from Word's UI alone.
6. **Review it** — clear; the Needs-Review banner is impossible to miss.
7. **Insert it** — works well via Workbench if built; noticeably worse via the compact InputBox fallback (Part 5/7) — this is the single most likely "this feels clunky" moment for a first-time user if they haven't set up the forms.
8. **Make a negotiation comment** — works, and is correctly protected against silently landing in the contract body (must select target wording first, Word.bas L1276-1279) — a good design a first-time user will appreciate once they understand why it insists on a selection.
9. **Save safely** — works well; the picker/permission flow is a little unusual (first Save requires an explicit Windows file-picker interaction even though you're "just clicking Save") but is clearly signposted by the status bar text, and I verified it behaves correctly under adversarial conditions (concurrent edits, permission loss).
10. **Recover from an error** — the crash-recovery confirm() prompt is plain-language and I verified it actually works; a first-time user would understand it without a manual.

**Net:** most of the friction a first-time lawyer would hit is in *setup and discoverability* (items 1, 3, 5, 7), not in the core edit/save/insert loop, which is already close to self-explanatory. This matches the audit's overall conclusion: fix onboarding and the Word-side front door before adding any new capability.

---

## PART 18 — FINAL REPORT

### A. EXECUTIVE VERDICT

| Dimension | Rating |
|---|---|
| Architecture | **GOOD** |
| Data safety (browser/HTML side) | **GOOD** — extensively adversarially tested and held up |
| Data safety (Word/VBA side) | **ACCEPTABLE** — sound design, genuinely unverified in the one runtime that matters |
| Word integration (capture/insert mechanics) | **ACCEPTABLE** |
| Word integration (discoverability: right-click, ribbon) | **NEEDS WORK** |
| Browser repository (HTML app) | **GOOD** |
| Installation | **NEEDS WORK** (fragile forms path is the default story; keyboard-only path is genuinely fine) |
| Daily usability | **ACCEPTABLE**, trending GOOD once onboarding/Ribbon gaps close |
| Maintainability | **GOOD** — plain files, no build step, no framework to rot |
| Scalability (to thousands of entries) | **GOOD**, verified to 2,000 |
| Overall product maturity | **ACCEPTABLE** |

**Would I personally start putting valuable precedents into this version? Yes, with conditions:**
1. Run the real-Word retest in `FOCUSED_WORD_RETEST.txt` yourself once, on synthetic text, before trusting rich (non-plain-text) capture of anything with tables or complex tracked changes — until then, prefer plain paragraph clauses for rich capture, or accept the plain-text fallback when prompted.
2. Set up a habit (calendar reminder, not app feature) of periodic JSON Backup until the rolling-backup addition (Part 12) exists.
3. If you ever intend to run more than one Clause Library file on this machine, re-run `ActivateCurrentClauseLibrary` every time you switch, and watch the "records skipped" status text.

None of these conditions is a reason to wait — they're the same "know your tools" caveats that would apply to any serious personal system, and the browser-side safety net (the part doing the actual, hard, adversarially-tested work of not losing or corrupting data) is genuinely solid.

### B. TOP FINDINGS
See Part 2 (§H-1 through §M-5) and Part 3's table for full detail, severity, evidence, and recommended fixes. Highest-value, in order: **H-1/H-2** (Track Changes correctness is real-Word-unverified and has no post-insertion equivalence check) → **H-4** (no discoverable Word-side front door) → **H-5** (single-generation rollback) → **H-3** (multi-library capture drift) → everything else in Part 2/3.

### C. WHAT IS ALREADY GOOD (do not rewrite)
- The single-file HTML-as-database architecture itself.
- The revision/commitId/savedAt optimistic-concurrency model and the disk-currency guard on Save (verified under attack).
- The crash-recovery system (verified under attack).
- The schema/version-migration logic in `normalizePayload`/`normalizeEntry`.
- The duplicate-ID detection-and-repair flow (verified under attack).
- The rich-formatting freshness (`richState`) model and its conservative default to plain text.
- The clean separation between "Word Comments" (negotiation comments, always requires a live selection) and clause-body insertion.
- The XSS-safe DOM-construction style and the deliberate `</script>` escaping on re-serialization.
- The honest, accurate self-documentation of what has and hasn't been tested — this is rare and should be preserved as a project norm, not seen as a weakness to paper over.

### D. SIMPLIFICATION OPPORTUNITIES (ranked)
1. Merge the two Word-connection buttons into one onboarding action.
2. Remember last-used category/family at capture time.
3. Demote Import to a secondary/overflow control.
4. Auto-run Health Check instead of requiring a manual click.
5. Ship the .dotm+forms pre-built; stop treating the VBIDE installer as the primary path.

### E. MISSING FUNCTIONALITY (ranked by lawyer value)
1. Real-Word verification pass (not new code — verification).
2. Rolling multi-generation backup.
3. Capture-side active-library guard.
4. Pre-built .dotm + minimal Ribbon.
5. Workbench double-click-insert / Ctrl+Enter-comment.
6. `status` → `tier` + `lifecycle` split.

### F. ARCHITECTURE DECISION
**KEEP CURRENT ARCHITECTURE.** Nothing surveyed in Part 10 is a clear improvement for this specific brief (one user, offline, Word-native, decades-durable, zero corporate friction). The identified problems are execution/verification and discoverability gaps, not architectural ones.

### G. IDEAL DAILY WORKFLOW
See Part 16 — already close to the current reality; the gap is onboarding and the Word-side front door, not the core loop.

### H. IDEAL PRODUCT SHAPE (v7)
Same two files (`.dotm` + `.html`), with: forms and a minimal Ribbon group pre-built into the .dotm (no installer needed for the common case); `status` split into `tier`/`lifecycle`; a rolling backup folder written alongside the HTML on every Save; a capture-side active-library guard; and a real-Word-verified Track Changes test pass covering tables and mid-revision boundaries. No new file formats, no new concepts, no database.

### I. PRIORITIZED ROADMAP

**MUST FIX BEFORE VALUABLE DATA**
- Real-Word retest of Track Changes capture (tables + mid-revision boundaries added to the test script).
- Rolling multi-generation backup.
- Capture-side active-library check.

**NEXT HIGH-VALUE IMPROVEMENTS**
- Pre-built .dotm + minimal Ribbon group.
- Remember last category/family at capture.
- Split status into tier/lifecycle.
- Workbench double-click-insert / Ctrl+Enter-comment.

**LATER**
- Unify the two search implementations (or formally accept and document the drift).
- Per-process-unique ID suffix.
- Onboarding "Connect to Word" wizard.

**DO NOT BUILD**
- Per-entry version history, word-level diff view, jurisdiction/agreement-type/usage-frequency fields, any workflow/approval/multi-user system, AI-assisted classification or drafting in the authoritative path, any cloud/sync service.

### J. FINAL ANSWERS

1. **Is the tool safe?** The browser/data-integrity layer is genuinely safe under adversarial testing. The Word-side Track Changes fix is soundly designed but not yet verified in the one environment that matters (real Word) — that is the one open safety question, and it is honestly disclosed as such by the vendor.
2. **Is the product too complicated?** No, structurally — but its *onboarding* asks a first-time user to understand more moving parts (cache vs. inbox vs. active-library vs. .dotm vs. forms installer) than the actual daily workflow requires. Simplify the front door, not the data model.
3. **Is anything important missing?** Yes — multi-generation backup and a discoverable Word-side entry point (Ribbon) are the two real gaps.
4. **Is anything currently unnecessary?** The VBIDE forms-installer-as-primary-path is more machinery than the problem needs, given a one-time manual .dotm build solves it more reliably.
5. **Is there a substantially better architecture?** No. Keep it.
6. **The five highest-value changes:** (1) real-Word Track-Changes verification including tables, (2) rolling backup, (3) capture-side active-library guard, (4) pre-built .dotm + minimal Ribbon, (5) split status into tier/lifecycle.
7. **Should you keep investing in this product?** Yes.
