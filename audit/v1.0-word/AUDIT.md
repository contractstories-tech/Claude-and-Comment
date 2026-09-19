# Clause Library 1.0 — independent product review and assurance audit

**Subject:** `ClauseLibrary-1.0.zip` (end-user package) and `ClauseLibrary-1.0-Source.zip` (source bundle)
**Date of audit:** 19 September 2026
**Auditor:** independent review, no prior involvement in the design

## Scope and honesty about the test environment

The audit environment is Linux. There is **no Microsoft Word, no Windows, no VBA runtime, and no
browser with local-file access**. Therefore:

- Every statement about Word object-model behaviour (`InsertXML`, `WordOpenXML`, `Range.Find`,
  Track Changes, numbering, tables, `Documents.Add`, `AddIns.Add`, `FileDialog`,
  `GetSetting`/`SaveSetting`, Startup-folder loading, UserForm rendering) is **reasoning from
  source**, not observation of a run.
- Installation, removal, macro-security behaviour and real browser rendering of `Browse.html`
  were **not** exercised.
- What *was* executed here: package integrity checks, VBA extraction and diffing against the
  supplied source, OOXML structure and metadata inspection, XML 1.0 conformance tests on the
  characters Word emits, a faithful re-implementation of the search algorithm run against
  realistic legal queries, and the generated catalogue's JavaScript run in a real DOM (jsdom)
  against adversarial clause text.

Findings are labelled **Observed** (verified here), **Likely** (strongly supported by the
implementation but not runtime-tested) or **Possible** (plausible, needs validation on Windows).

---

## 1. How the product actually works

### 1.1 Shape of the system

Two shipped binaries and three text files. No installer, no service, no network, no runtime
dependency beyond Windows and desktop Word.

- **`Start Here.docm`** — a setup document carrying a small VBA project (`ClauseLibrarySetup`)
  with `CLPlatform`, `CLStore`, `CLSetup` and a 6-line `SetupSupport` shim, plus a "Clause
  Library Setup" ribbon tab. It is the installer and the uninstaller.
- **`ClauseLibraryPersonal.dotm`** — the product. VBA project `ClauseLibraryProduct`:
  `CLActions` (workflow), `CLStore` (storage), `CLPlatform` (file/crypto/XML primitives),
  `CCLRichContent` (the Word formatting engine, source file `CLRich.bas`), `CLRecovery`
  (history and backup), `CLCatalogue` (HTML export), and two UserForms (`frmLibrary`,
  `frmHistory`). Its document body is empty, so it injects no content.

### 1.2 Installation

`CLSetUp` copies the `.dotm` into Word's Startup folder, verifies the copy byte-for-byte with
SHA-256, writes an `.owner` marker beside it, registers the path under
`HKCU\…\VB and VBA Program Settings\ClauseLibraryPersonal`, creates the library at
`%LOCALAPPDATA%\ClauseLibraryPersonal` if absent, keeps a second copy of both files in
`%LOCALAPPDATA%\ClauseLibrarySupport` so removal is possible later, and finally calls
`AddIns.Add … Install:=True` so the ribbon appears without restarting Word. A failure after the
copy triggers a rollback. Nothing touches `Normal.dotm`, trust settings or policy.

### 1.3 Where the authoritative data lives

One folder — the "library root" — containing:

```
library.xml                 identity marker: product, schema="1", id=<guid>
entries/<guid>.xml          one file per clause or comment  ← the authoritative record
history/<guid>-r<10 digits>.xml   an immutable full copy of every superseded revision
writer.lock                 a zero-length file used purely as a cross-process mutex
Browse.html                 a generated read-only snapshot (optional)
```

The root's path **and** its `library.xml` id are stored together in one registry value
(`id|path`). On every access `CLRoot()` re-reads both and refuses to proceed if they disagree —
so the tool will not silently adopt a different folder or create an empty replacement.

An entry file is:

```xml
<entry product="ClauseLibraryPersonal" schema="1" revision="N" id="…" digest="…">
  <data>
    <kind/><title/><plain/><rich/><notes/><topic/><tags/><family/><role/>
    <applicability/><state/><organised/><favourite/><created/><updated/><source/><comment/>
  </data>
</entry>
```

`rich` holds a **complete Flat OPC Word package** (`Range.WordOpenXML`) as XML-escaped text.
`plain` holds the same wording as flat text and is what search and preview use. `digest` is
SHA-256 over `id + revision + the serialised <data> subtree`.

### 1.4 Writing

`CLSave` is the single write path:

1. take an exclusive lock on `writer.lock` (`Open … Lock Read Write`) — serialises writers
   across Word processes;
2. load the current file and compare its revision to the caller's expected revision — a stale
   editor is refused;
3. copy the current record into `history/` and read it back to verify;
4. bump the revision, stamp `updated`, recompute the digest;
5. re-parse the serialised entry and re-check the digest before committing;
6. write to `<file>.<guid>.pending`, read it back and compare, then `MoveFileExW` with
   `REPLACE_EXISTING | WRITE_THROUGH`;
7. read the committed file back and re-validate it.

That is five independent integrity mechanisms layered on one save. More on that in §11.

### 1.5 Capture

`CLCaptureSelection` duplicates the selection, refuses partial tables, prompts once if the
selection carries tracked changes, then serialises the range to Flat OPC, opens a **hidden
disposable document**, re-inserts the XML there, and in that copy accepts all revisions, deletes
all comments, deletes all hyperlinks, unlinks all fields, rejects non-picture inline shapes and
any floating shapes, deletes hidden-formatted text, and strips every `w:sectPr`. The result is
re-validated against a blocklist (`w:ins`, `w:del`, `w:comment*`, `w:fldChar`, `w:altChunk`,
`w:object`, `w:dataBinding`, external relationships, `vbaProject`/`embeddings`/`customXml`
parts) before it is stored. **The user's document is never modified.**

Any comments attached to the captured range are copied into the entry's `notes` field.

### 1.6 Retrieval

`frmLibrary` calls `CLList()`, which reads, parses, digest-verifies **every** entry file, then
filters in memory with `CLMatches` — an AND of case-insensitive substring matches across nine
fields. The list shows two columns: title and kind. Selecting an item shows a plain-text preview
and an editable details panel.

### 1.7 Reuse

`CLInsertEntry` reloads the entry, refuses non-Active items, confirms replacement of a selection,
warns if the wording contains `[` or `{{`, then either sets `Range.Text` (plain path, wrapped in
a custom Undo record) or calls `CCLInsertRichDevelopment`, which validates the payload, refuses
protected/read-only targets and partial-table positions, rehearses the insertion twice in hidden
documents, then inserts inside a custom Undo record and rolls back on failure. Comment entries
take a separate path that adds a real Word comment after an explicit confirmation showing the
text.

### 1.8 Editing, history, backup, removal

"Edit wording in Word" opens a new unsaved document carrying three document variables (entry id,
expected revision, library root); "Save wording" re-captures it and commits a new revision.
"Previous versions" lists the history files and can restore one **as a new revision**. "Back up
library" takes the writer lock and makes a fully re-verified copy of every XML file into a
timestamped folder, writing `backup-complete.xml` last. "Restore a backup" validates the backup,
copies it to a new `ClauseLibraryRestored-<guid>` folder and switches the connection to it.
"Remove Word integration" un-installs and deletes the Startup template only.

---

## 2. Overall product and trust assessment

**The central question: can a commercial lawyer rely on this for years with a valuable
repository?**

**Today: not yet — but the gap is narrower than it looks, and it is not mainly a correctness gap.**

What this product gets genuinely right is *unusual*. The source document is never mutated. The
capture pipeline is paranoid in the right places. Tracked changes are resolved in a disposable
copy rather than in the lawyer's draft. Private notes are structurally separated from insertable
content and that separation is enforced at the code level, not by convention. The identity
marker means the tool will refuse to invent an empty library rather than silently lose one. The
package is honest — it does not claim signing it does not have, and `VALIDATION.md` states its
limits plainly. I verified the shipped VBA is byte-equivalent to the supplied source (modulo
VBA's own identifier-case normalisation), which is more provenance than most tools of this kind
offer.

What stops me recommending reliance today is a cluster of three things:

1. **A storage design that does not survive its own success.** Each entry costs roughly 60–100 KB
   regardless of how short the clause is, because a full Flat OPC package (including a ~42 KB
   styles part and an ~8.7 KB theme part) is escaped and nested inside the entry XML. Every
   library refresh re-reads, re-parses and re-hashes *all* of it. See §3.2 — this is the finding
   with the largest long-term consequence.
2. **Realistic content that cannot be captured at all, failing with an opaque message.** Footnote
   marks, inline images, page breaks and non-breaking/optional hyphens all produce characters
   that XML 1.0 forbids, and only three of them are sanitised. See §3.1.
3. **A daily workflow that will quietly erode the repository.** Auto-titles are the first 90
   characters of the clause; there is no sort, no date, no duplicate detection, no usage signal,
   and the information shown at the moment of reuse is a six-line plain-text box. See §5 and §6.

None of these is a design dead end. All three have proportionate fixes.

---

## 3. Reliability and data-integrity assessment

### 3.1 Realistic Word content silently cannot be stored — **Observed / Likely**

`CCLPlainWording` normalises exactly three characters: `Chr(7)`→tab, `Chr(11)`→LF, CR→LF.
Everything else Word's `Range.Text` can emit passes through into the `plain`, `comment` and
`notes` fields, which are then written as XML text nodes.

I tested every such character against a conformant XML 1.0 parser, in both raw and numeric-
character-reference form:

| Char | Meaning in Word text | Handled? | Legal in XML 1.0? |
|------|---------------------|----------|-------------------|
| 1  | inline picture / drawing anchor | no | **rejected** |
| 2  | footnote / endnote reference mark | no | **rejected** |
| 5  | comment (annotation) reference mark | no | **rejected** |
| 7  | table cell / row end | yes → tab | n/a |
| 11 | manual line break (Shift+Enter) | yes → LF | n/a |
| 12 | page break / section break | no | **rejected** |
| 13 | paragraph mark | yes → LF | n/a |
| 14 | column break | no | **rejected** |
| 19 / 21 | field begin / end | no | **rejected** |
| 30 | non-breaking hyphen (Ctrl+Shift+-) | no | **rejected** |
| 31 | optional (soft) hyphen | no | **rejected** |

Numeric references do not help: `&#1;` is as illegal as a raw `Chr(1)`.

Consequence (**Likely**, would fail at `CLSave`'s pre-commit re-parse or earlier): capturing a
clause that contains a footnote reference, a non-breaking hyphen, an optional hyphen, a page
break, or an inline image fails with *"Capture was not completed. This library file is incomplete
or unreadable. It has been preserved."* — a message about *library files* shown to a user who was
capturing *text*, with no indication of what to remove or that the problem is a character.

This is sharpened by an internal contradiction: `CLRich.bas:58-60` **deliberately permits**
`wdInlineShapePicture` in a capture, and that is precisely the case that then produces `Chr(1)`
and breaks the save. The one media type the rich engine allows is one the storage layer rejects.

Non-breaking and optional hyphens are not exotic — they appear routinely in professionally
typeset contracts. Footnotes appear in most finance and M&A drafts.

**Fix is small:** strip or substitute all characters outside the XML 1.0 white-list in
`CCLPlainWording` (`Chr(1)`→"", `Chr(2)`→"", `Chr(30)`→"-", `Chr(31)`→"", `Chr(12)`/`Chr(14)`→LF,
and a catch-all for anything else below `Chr(32)` other than tab/LF). Half a dozen lines.

### 3.2 Storage cost per entry is dominated by fixed overhead — **Likely**

`Range.WordOpenXML` returns a *whole* Flat OPC package, not just the selected runs. Using the
shipped template's own parts as a measure of what Word emits:

```
word/styles.xml        42,760 bytes
word/theme/theme1.xml   8,717 bytes
word/settings.xml       3,287 bytes
word/document.xml       2,940 bytes
word/fontTable.xml      1,749 bytes
word/webSettings.xml    1,069 bytes
numbering.xml (added by any numbered clause)  ~4,000 bytes
                        ──────────────────
fixed overhead         ~64,500 bytes per entry
```

That package is then stored as **XML-escaped text inside the entry envelope** (`CLSet e, "rich",
xml`), inflating it by a further ~20 %. So a one-sentence indemnity clause of 145 characters
becomes an entry file of roughly **76 KB**, of which under 0.2 % is the lawyer's wording.

Projected, assuming eight edits over a clause's life (history keeps a full copy each time):

| clauses | entries | history | total library | re-read + re-hashed on every refresh |
|--------:|--------:|--------:|--------------:|-------------------------------------:|
| 100     | 7.8 MB  | 62.6 MB | 70 MB         | 7.8 MB |
| 300     | 23.5 MB | 187.8 MB| 211 MB        | 23.5 MB |
| 1,000   | 78.3 MB | 626 MB  | 704 MB        | 78.3 MB |
| 3,000   | 235 MB  | 1.88 GB | 2.1 GB        | 235 MB |

**The author can confirm the per-entry figure in ten seconds**: select one short clause in Word
and run `?Len(Selection.WordOpenXML)` in the Immediate window. If it returns tens of thousands,
this finding holds.

### 3.3 The refresh cost is paid on every trivial action — **Observed (code) / Likely (impact)**

`RefreshEntries` → `CLList()` reads, UTF-8-decodes, DOM-parses, re-serialises the `<data>`
subtree, re-encodes it and SHA-256s it, for **every entry file**. It is called on: opening the
library, every capture, every "Save details", every Archive, every Trash, every Restore, every
"Save wording", and every Refresh.

So ticking the **Favourite** checkbox and pressing Save on a 300-entry library re-reads and
re-hashes ~23 MB through ADODB.Stream + MSXML + CryptoAPI on the UI thread, with no progress
indication and a modeless form that will appear frozen. `frmLibrary` also holds every entry as a
live MSXML DOM for as long as it is open — MSXML DOM overhead is typically 5–10× the text size,
so a few hundred entries is plausibly several hundred MB resident. **On 32-bit Word (still common
in managed corporate estates) that is a hard ceiling, not a slowdown.**

The architectural cause is simple and the fix is simple: **the rich payload should not live inside
the record that listing and searching must read.**

### 3.4 Backup refuses to run when you most need it — **Observed**

`CLBackupTo` (`CLRecovery.bas:41`) calls `CLLoadPath` on every source file with **no error
isolation**. One corrupt or hand-edited entry aborts the entire backup: *"Backup did not
complete."* — and nothing is backed up, including the 99 % of entries that are perfectly healthy.

`CLRestoreBackupTo` has the same all-or-nothing property.

This is the interaction of two individually defensible decisions ("verify everything", "fail
closed") producing exactly the wrong outcome: the moment your library develops a problem is the
moment you most want a copy of it. Note the inconsistency — `CLList` *does* isolate bad entries
and reports them while showing the rest. Backup should behave the same way: copy what is healthy,
copy the damaged files verbatim without validating them, and report both counts.

### 3.5 A corrupt entry has no recovery path in the product — **Observed**

When `CLLoadPath` rejects an entry, the message says *"Entry integrity check failed. **Use its
history snapshot**; the file was preserved."* But:

- `CLList` excludes the entry, so it never appears in the library list;
- `cmdHistory_Click` requires a selected list item;
- therefore the History window — the advised remedy — **cannot be reached for that entry**.

The only routes back are Explorer plus manual XML surgery, or a full backup restore. For a tool
whose audience is explicitly non-technical, an error message whose advice is unreachable through
the UI is a real defect. A "Problem entries…" button next to the existing
*"Some entries could not be read"* status, offering *Restore latest good version from history*,
would close it.

### 3.6 The digest is more likely to brick an entry than to save one — **Observed / Possible**

What does `digest` uniquely catch that the rest of the stack doesn't?

- Truncated or partial writes → already prevented by `CLAtomicWrite` + `MoveFileEx`, and a
  truncated XML file would not parse anyway.
- Disk bit-rot → real, but rare on modern NTFS, and the history copy has the same exposure.
- Tampering → not a security control; it is unkeyed and trivially recomputed.

What it *does* reliably catch is a user hand-editing an entry file — and the product actively
invites that by putting a **"Library folder"** button in the UI that opens Explorer on the data.
After such an edit the entry is permanently unreadable through the product, with no
"recompute digest" or "accept this file" path (§3.5).

Separately (**Possible**): the digest is computed over MSXML's *re-serialisation* of the `<data>`
subtree, not the file's bytes. Any difference in how a future MSXML build escapes or orders
content would invalidate every entry at once. `CLFileSha` already exists in the same module and
hashes actual bytes — the product carries two hashing mechanisms and uses the more fragile one
for the data that matters.

### 3.7 Smaller data-integrity items — **Observed**

- **Orphaned `.pending` files.** If `MoveFileExW` fails after `CLWrite pending` succeeds, the
  `.pending` file stays in `entries/` forever. It is correctly ignored by `CLList` (extension
  filter) but nothing ever cleans it up.
- **`CLStamp` calls `Now` twice** (`Format$(Now,"yyyy-mm-dd") & "T" & Format$(Now,"hh:nn:ss")`),
  so a timestamp crossing a second/midnight boundary can be internally inconsistent. It also uses
  `":"`, which `Format$` treats as the **locale time separator** — on a machine configured with
  `.` the timestamps become `2026-09-19T14.05.33`. Neither breaks anything today; both make the
  stored timestamps less trustworthy than they look.
- **`CLHistory` has no error isolation** (unlike `CLList`): one bad history file makes the entire
  version list for that entry unavailable.
- **Nothing is ever sorted.** `CLList` and `CLHistory` return files in `FileSystemObject`
  enumeration order. Since entry filenames are GUIDs, the library list order is effectively
  arbitrary and can change between sessions.
- **Nothing is ever pruned.** No trash purge, no history cap, no backup rotation. Growth is
  monotonic by design, and trashed items keep costing full read+hash on every refresh.

---

## 4. Microsoft Word and drafting-workflow assessment

*Nothing in this section could be executed. It is source reasoning plus knowledge of Word's
documented behaviour.*

### 4.1 What the rich engine gets right — keep it

- Capture into a **disposable hidden copy**; the lawyer's document is never touched. This is the
  single most important safety property in the product and it is implemented correctly.
- **Stripping `w:sectPr`** so an inserted clause cannot change the destination's page setup,
  margins or headers. Non-obvious, and exactly right.
- **Explicit refusal over silent flattening** for embedded objects, floating shapes and partial
  tables.
- **A blocklist validator** run on the payload *after* cleaning and *again* before insertion,
  covering revisions, comments, fields, `altChunk`, data bindings, external relationships and
  private parts. Well chosen.
- Insertion honours the destination's Track Changes state and wraps the change in a custom Undo
  record with rollback on failure.

### 4.2 The whole-document rehearsal should be removed — **Observed (code) / Likely (cost)**

`CCLInsertRichDevelopment` performs two rehearsals. The first (lines 127–130) inserts the payload
into a fresh hidden document and checks its text matches the stored `plain` — cheap and useful.

The second (lines 131–135) does this:

```vba
beforeXml = doc.Content.WordOpenXML          ' serialise the ENTIRE destination contract
Set probe = Documents.Add(Visible:=False)
probe.Content.InsertXML beforeXml            ' rebuild the ENTIRE contract in a hidden document
probe.Range(target.Start, target.End).InsertXML xml
probe.Close wdDoNotSaveChanges               ' …and throw it away
```

Three problems:

1. **Its result is never checked.** Nothing is asserted about `probe` afterwards. The only signal
   is "did it throw" — which the real insertion path would also produce, and which the custom
   Undo record already rolls back.
2. **Its position assumption is unsound.** Character offsets in `probe` are not guaranteed to
   match `doc` after a Flat OPC round-trip (list numbering, field results and the final paragraph
   mark can shift them). So the rehearsal may be inserting somewhere else entirely, making its
   "did it throw" signal unreliable in *both* directions.
3. **It is the dominant cost of every insertion.** For a 150-page negotiated contract with
   schedules and images, `Content.WordOpenXML` can run to tens of megabytes of XML held in a VBA
   `String`, immediately re-parsed into a second full Document object. The error path at line 151
   then serialises the whole document *again*. On 32-bit Word this is a plausible out-of-memory
   condition; on 64-bit it is seconds to tens of seconds of apparent hang, on the most frequent
   action in the product.

This is the clearest instance in the product of a technically sophisticated mechanism whose cost
is certain and whose benefit is speculative. **Delete lines 131–135 and the `beforeXml` logic in
the error handler. Keep the payload probe.**

### 4.3 The exact-text gate is a long-term time bomb — **Likely**

Line 129 refuses insertion unless `CCLPlainWording(probe.Content.Text)` is **character-identical**
to the `plain` string frozen at capture time. `plain` was produced by whatever Word build was
installed then; the comparison value is produced by whatever Word build is installed now.

If a future Word update changes any normalisation in the Flat OPC round trip — smart quotes,
`w:noProof`, soft hyphens, list separators — then **every rich insertion for every entry in the
library fails simultaneously**, permanently, with *"The formatted wording does not match this
saved entry."* There is no override, no diagnostic showing what differs, and the message blames
the entry rather than explaining that a safety check fired.

For a repository intended to outlive several Office versions, an unconditional equality gate
against a value frozen years earlier is the wrong shape. Make it a **warning with an
"Insert anyway" option**, or compare after normalising whitespace and quote characters.

### 4.4 Cross-references are silently frozen — **Observed (code) / Likely (consequence)**

`CLRich.bas:57` unlinks every field in the capture copy. A clause reading *"as set out in clause
14.3"* via a `REF` field becomes the literal text *"14.3"*. Inserted into a different contract
with different numbering, that cross-reference is **wrong and looks right**. Nothing warns the
user; `READ_ME.txt` and `VALIDATION.md` do not mention cross-references at all.

This is a drafting-risk item, not a technical one, and it deserves a line in the documentation at
minimum and ideally a capture-time notice ("this wording contained N cross-references; they have
been fixed as text").

### 4.5 Style and numbering import — **Likely**

Because the payload is a full Flat OPC package including `styles.xml` and `numbering.xml`,
`InsertXML` brings the captured document's style *definitions* into the destination. Word's
standard behaviour on a name collision with differing definitions is to create a variant
(`Body Text1`, `ListParagraph1`). Over dozens of insertions from differently-styled source
documents, the destination contract accumulates duplicate styles that survive into the delivered
document. `VALIDATION.md` gestures at this ("conflicting custom styles still require checking")
but frames it as a visual issue; the real consequence is style-list pollution in a document that
goes to the other side.

### 4.6 Hidden-text removal has no loop guard — **Possible**

`CLRich.bas:62-72` repeatedly searches from the start of the temp document for hidden-formatted
text and deletes it. If `Find.Execute` ever returns `True` on a range that `Delete` cannot remove,
the loop re-searches from the same position and **never terminates** — Word hangs with no user
escape, during a capture. Whether that state is reachable is Word-behaviour-dependent and I
cannot test it. A maximum-iteration counter is two lines and removes the class of risk entirely.

Separately: deleting hidden text **silently** contradicts the module's own stated principle
("explicit refusal is preferable to silently losing unsupported content"). Hidden text in legal
drafts is sometimes deliberate.

### 4.7 Table cells cannot receive rich wording at all — **Observed**

Insertion inside a table cell is refused outright. Pricing schedules, SLA tables and DPA annexes
are ordinary contract furniture, so this is a real capability gap rather than only a safety
measure. The plain-text path *does* work in a cell — but the refusal message
("Place the cursor outside the table…") doesn't say so.

### 4.8 "Optional deliberate tracking" does not exist — **Observed**

`PRODUCT.md` promises *"Insert using current Word tracking state, with optional deliberate
tracking."* In production, `CLActions.bas:73` passes `target.Document.TrackRevisions` as
`trackInsertion`, so `CCLInsertRichDevelopment` sets Track Changes to the value it already has.
The parameter is exercised only by the test suite. There is no UI for it. Either build it or
remove the claim.

---

## 5. Product usefulness and day-to-day legal workflow

This is where I am most concerned, and it has nothing to do with correctness.

### 5.1 Reuse is too many steps for the thing you do most — **Observed**

To insert a favourite clause: click the Clause Library tab → Library → wait for the full library
to load and verify → type a search → click the item → click Insert → confirm the replacement →
confirm the placeholder warning. There is **no quick-insert on the ribbon, no favourites gallery,
no recent list, and no keyboard shortcut**. During a live negotiation, that is the difference
between a tool you reach for and a tool you stop opening.

Word's own Building Blocks gallery — one click from the ribbon, previewed, inserted — is the bar
this has to clear, and today it doesn't.

### 5.2 The information at the point of reuse is insufficient — **Observed**

The list has two visible columns: a truncated title and "Clause" or "Clause / New". The preview
is a **fixed 458 × 104 point plain-text box** (about six lines) — and `Range.Text` **omits list
numbers**, so a clause stored as a three-level numbered list previews as unnumbered running text
that does not resemble what will be inserted. Nothing shows the formatting, the table structure,
or the numbering.

Faced with twelve indemnity clauses, the lawyer has a truncated first line and six lines of
unnumbered text with which to choose. That is not enough for a drafting decision, and it is the
question the product exists to answer.

### 5.3 Data the product already collects and never shows — **Observed**

I checked every stored field against every read site:

| field | written | ever read? |
|---|---|---|
| `source` (originating document name) | on every capture | **never** — not displayed, not searched |
| `created` | on every capture | **never** |
| `audience` | on every capture (`"Private"`) | **never** |
| `updated` | on every save | only in the history window's version list |

So the library **knows** when each clause was captured, when it was last changed and which
document it came from, and shows the user **none of it**. There is no date column, no sort, and
no way to answer "which clauses have I added recently" or "where did this come from".

This is the cheapest high-value improvement available: surface `created`/`updated`/`source`,
add a sortable date column, and add `source` to the search field list. Perhaps thirty lines.

### 5.4 Titles will make the library unusable at scale — **Observed**

`CLActions.bas:44` sets the title to the **first 90 characters of the clause**. Capture is one
click with no prompt, and organising is explicitly optional. The predictable steady state after a
year is three hundred entries titled *"The Supplier shall indemnify and hold harmless the
Customer against all Losses arising…"*, *"The Supplier shall indemnify the Customer against…"*,
*"The Supplier shall indemnify…"* — visually indistinguishable in a 235-point column.

The zero-friction capture philosophy is right. The auto-title implementation undermines it.
A better default costs little: title from the **nearest preceding heading** in the source
document, falling back to the first sentence; or a single-field "Name this clause?" prompt with
the auto-title pre-filled, dismissible with Enter.

### 5.5 No duplicate detection — **Observed**

Re-capturing the same clause from a later draft (the most natural thing a lawyer does) silently
creates a second entry with a near-identical title and no link to the first. Over years this is
the principal degradation mechanism for a precedent repository. A hash of the normalised `plain`
text, checked at capture, would let the tool say *"You already have very similar wording — add
as a new variant of that family, or keep separate?"*

### 5.6 The placeholder dialog will train users to ignore all dialogs — **Observed**

`CLActions.bas:64` warns whenever the wording contains `[`. Square brackets are ubiquitous in
real contract text: `[Insert date]`, `[2026]`, statutory citations `[2019] EWCA Civ 1`, defined
term markers, drafting options. This dialog will fire on a large proportion of insertions.

The cost is not the click. The cost is that it sits next to two dialogs that genuinely matter —
*"Replace the selected wording in <document>?"* and *"Add this comment… it will be visible to
anyone receiving the document"* — and teaches the user to dismiss all three reflexively. An
always-on warning is a warning that has been turned off.

Better: don't warn. Insert, then move the cursor to the first placeholder, or apply a highlight.
The information is more useful *after* insertion than before it.

### 5.7 Success feedback is too quiet — **Observed**

Capture, insertion, "Save wording" and comment capture all report success only via
`Application.StatusBar`. The status bar is easy to miss, is overwritten by other Word activity,
and is not reset — so a stale *"Saved to Clause Library"* can linger and be mistaken for a fresh
confirmation. With the library form closed, a lawyer capturing five clauses in a row gets
essentially no evidence that anything happened.

### 5.8 Would this become a habit?

Capture: yes — one click, no questionnaire, source untouched. That part is genuinely well judged.
Retrieval and reuse: probably not, for the reasons in §5.1–5.4. The likely trajectory is
enthusiastic capture for a few weeks, followed by the list becoming unnavigable, followed by
quietly not opening it — which is the specific failure mode the brief asks about.

---

## 6. UX and usability

- **The window cannot be resized.** VBA UserForms are fixed-size; the layout is absolute
  coordinates at `Zoom = 90`. On a large monitor the library is a small fixed panel; the preview
  stays six lines regardless of available space.
- **The layout depends on `Zoom = 90` to fit** — `lblStatus` sits at Top 534 in a form of Height
  536. At 90 % it renders inside the client area. It works, but it is fragile against any future
  layout change. (**Possible**; rendering not testable here.)
- **Quoted phrase search returns nothing.** Executed: `CLMatches` splits on literal spaces, so
  `"force majeure"` becomes the tokens `"force` and `majeure"`, neither of which matches. A
  lawyer typing quotes — near-universal search behaviour — concludes the clause isn't there.
  Stripping `"` from the query is a one-line fix.
- **Substring matching without word boundaries generates noise.** Executed against a realistic
  clause set: `ip` returns *Intellectual property licence*, *Principal contractor* **and**
  *Recipient obligations*; `act` matches *contractor*; `cap` matches *uncapped*. Short legal
  terms are exactly the ones lawyers search for.
- **Zero results offers no help** — just "0 found". No "try fewer words", no fallback to OR.
- **No sort control** and, as noted, arbitrary list order.
- **"Save wording" is always enabled.** Pressing it on an ordinary document produces
  *"The library was not updated. The requested member of the collection does not exist."* — a raw
  VBA collection error (5941) surfaced verbatim to a lawyer. A `getEnabled` ribbon callback, or a
  guarded read of the document variable, fixes it.
- **Writer-lock contention surfaces as "Permission denied"** — VBA error 70 passed straight
  through `CLSave`'s handler. Reachable whenever two Word instances are open. The friendly
  message this deserves ("Another Word window is saving to your library; try again in a moment")
  is not there.
- **After restoring a version, the library form is not refreshed** — it says "Refresh the library
  to see it", but until you do, the preview shows the *superseded* wording as current, and a
  subsequent "Save details" fails the stale-revision check. Insertion is safe (it reloads), but
  this is stale information presented as current.
- **The draft document is visually indistinguishable from a contract.** "Edit wording in Word"
  opens an ordinary unsaved document with no watermark, no distinctive name, no header. With six
  documents open there is nothing to tell you which one is the library draft — and nothing stops
  you inserting a clause *into* the draft and then saving that as the entry.

**Good UX decisions worth preserving:** the "Not organised" reminder that never gates reuse; the
explicit "this comment will be visible to anyone receiving the document" confirmation with the
text shown; the tooltip on the notes field stating it is never inserted *and* that it is included
in backups; `cmdClose.Cancel = True` so Escape closes the form; the unsaved-details prompt on
every navigation path (`LeaveCurrent` is called from list click, edit, state change, refresh,
history and form close — consistently, with no gaps I could find).

---

## 7. Backup, recovery and long-term continuity

### What is right

- Backup takes the **writer lock**, so it cannot copy a half-written record.
- `backup-complete.xml` is written **last** — an interrupted backup can never be mistaken for a
  complete one. The same "marker last" discipline is used in restore. This is textbook and
  correct.
- Restore creates a **separate** library and leaves the original untouched.
- Every copied file is validated on both sides.
- The backup dialog explicitly says the backup contains private notes and recommends another
  drive. Honest and appropriate.

### What is wrong

- **§3.4: one bad file aborts the whole backup.** This is the most serious finding in this
  section.
- **No rotation, no pruning, no incremental.** Each backup is a full copy into a new timestamped
  folder. With the growth profile in §3.2, weekly backups of a 300-clause library are ~210 MB
  each; three years of them is ~33 GB, and the user must manage that entirely by hand.
- **Restore produces a new folder name every time.** After one restore the live library is
  `%LOCALAPPDATA%\ClauseLibraryRestored-<guid>`, after two there are two such folders, and
  `READ_ME.txt`'s statement that the data lives in `ClauseLibraryPersonal` is no longer true.
  Nothing in the UI shows which folder is currently connected.
- **A restored library keeps the original's identity.** `CLRestoreBackupTo` copies `library.xml`
  verbatim, so the restored copy has the *same* id as the live one. The identity marker therefore
  protects against "this is not your library" but **not** against "this is an older copy of your
  library". A user who restores a backup and later uses "Locate library" to pick the wrong folder
  gets no warning and silently works from stale precedent. Adding a `restoredFrom`/`created`
  attribute and showing the connected folder and its date in the form's status line would close
  this.
- **No export.** The only way to read the library outside this tool is `Browse.html`, which is
  plain text, active items only, and excludes private notes. There is no "export everything to a
  folder of .docx files" and no documented file format for recipients. For a repository a lawyer
  is asked to trust for years, *"can I get my content out if this stops working?"* is a trust
  question, and today the honest answer is "partially".
- **No complete uninstall.** `CLRemoveIntegration` deletes only the Startup template and its
  registry key. `%LOCALAPPDATA%\ClauseLibrarySupport` (containing full copies of both product
  files) and the `Library\Connection` registry value remain indefinitely, with no documented way
  to remove them. For a confidential local repository on a machine that may be returned to an
  employer, "how do I remove this completely" has no answer.

---

## 8. Architecture and maintainability

### Sound decisions

- **Word is the sole writer.** No second editable store, no synchronisation problem. Correct.
- **One file per entry.** No global master file to corrupt; a damaged entry is one damaged entry.
- **The catalogue is explicitly derived and dispensable.** Correct instinct.
- **The identity marker** preventing silent adoption of a wrong folder.
- **Byte-verified template installation** with rollback.
- **Provenance is verifiable** — I diffed the VBA extracted from both shipped binaries against
  the supplied `.bas` files and they are identical apart from VBA's own identifier-case
  normalisation (`f.Name`/`f.name`, `.XML`/`.xml`). Shipped metadata is scrubbed; the author's
  username appears nowhere in the binaries. Both are valid OOXML packages with
  `[Content_Types].xml` first. The manifest checksums verify. This is better release hygiene than
  most tools of this class.

### The structural problems

**The rich payload lives inside the record.** Everything in §3.2/§3.3 follows from this one
decision. Storing the payload as a sibling `.docx` file — entry XML holds metadata and a
filename, payload is opened only on preview/insert — would leave entry files at 1–2 KB, make
listing and search effectively instant, make the library directly inspectable (double-click the
`.docx`), and let insertion use `Range.InsertFile`, which is Word's oldest and best-tested
insertion path. This is the single highest-leverage architectural change available.

**The storage layer is duplicated across two VBA projects.** `CLPlatform`, `CLStore` and
`CLSetup` exist as independent copies in the `.dotm` and the `.docm` (I verified they are
currently identical apart from case). Any future fix must be applied to both, and nothing checks
they agree at runtime.

**The build is a three-stage code-generation pipeline that cannot be run as documented.**
`make-builder.cjs` emits `ProductBuilder.bas`; `make-install-builder.cjs` transforms that by a
chain of exact-substring `.replace()` calls (including CRLF-sensitive ones, only one of which is
verified afterwards); the result is imported into Word and run with "Trust access to the VBA
project object model" enabled; then a Python + lxml script patches the ribbons. And
`make-builder.cjs:49` hard-codes:

```
root = "C:\Users\harsh\Documents\Codex\2026-09-16\files-mentioned-by-the-user-clause\work\product"
```

So `BUILD.md`'s instruction to "put these directories under `work/product/` in a local checkout"
**does not work** for anyone else, or for the author on a new machine, without editing a path
buried inside a generated string. (It also leaks the author's Windows username in a bundle
described as shareable — cosmetic, but it is there.)

This machinery exists almost entirely to place UserForm controls at absolute coordinates — a
one-time job. In two years, fixing a one-line bug in `CLStore` should not require reconstructing
this chain. **Build the forms once, keep the `.dotm` as the maintained artefact, export the
`.bas` files for version control, and delete the generators.**

**`validate-sources.cjs` mutates production source.** Its first action is to read
`word/CLCatalogue.bas`, strip duplicate `Public Sub CLRibbonBrowse(` lines, and **write the file
back**. A script named "validate" that silently edits production source to paper over a generator
defect is a build-quality red flag: it means the generator can emit a duplicate definition and
the response was to delete the symptom.

**No product version exists anywhere** — not in the VBA, not in `library.xml` (only
`schema="1"`), not in the ribbon, not in the registry. A user cannot tell which build they have;
a future release cannot tell which build created a library; support has nothing to key on. This
should be fixed before, not after, a second version exists.

---

## 9. Security and confidentiality

The threat model is right for the product: a single lawyer's machine, no network, no accounts,
no server. Judged against that:

**Good**

- **No network egress anywhere.** No HTTP calls, no telemetry, no external references. I checked
  the generated catalogue for external resource references and found none.
- **Private notes are structurally separated.** `CLInsertComment` refuses to insert the `notes`
  field; the rich payload is built from a cleaned copy that never contains it; `CLBrowse` emits
  only `title`, `kind`, `topic`, `tags`, `applicability` and `plain`. This is enforced in code,
  not by convention.
- **`ProhibitDTD`, `resolveExternals = False`** on every DOM — XXE is closed off.
- **The validator rejects `vbaProject`, `embeddings/`, `customXml/` parts and external
  relationships** in a payload, so a clause captured from a hostile document cannot carry active
  content into the library or back out into a contract.
- **HTML escaping holds.** I generated `Browse.html` with adversarial clause text
  (`</pre><script>`, `</script><script>`, `onerror=`, backticks, `${}`) and parsed the result:
  exactly one `<script>` element, no breakout, no injection. Archived and trashed wording is
  correctly excluded. A CSP meta tag is present with `default-src 'none'` and `connect-src
  'none'`.
- **The package is honest about being unsigned.** I confirmed there is no digital signature
  stream in either `vbaProject.bin`, matching what `VALIDATION.md` says. The VBA project is also
  not password-locked, which is the right choice for an auditable personal tool.
- **No trust settings, `Normal.dotm` changes or policy modifications**, as claimed.

**Concerns**

1. **Source-document comments are silently absorbed into the library.** `CLActions.bas:34-38`
   copies every comment attached to the captured range into the entry's `notes` field. This is
   deliberate and tested (`CLWorkflowTests:25`, labelled *"Private capture rationale"*), and the
   test's framing assumes the comments are the *user's own*. In practice a lawyer most often
   captures from a **received or marked-up draft**, where the comments are the counterparty's,
   another firm's, or internal privileged commentary. Those comments then become a permanent,
   searchable part of the precedent repository, are copied into every backup, and are shown in a
   panel labelled "Private guidance" — with **no indication at capture time that this happened**
   and no opportunity to decline. This is a defensible feature with an undisclosed confidentiality
   consequence. At minimum, tell the user ("3 comments from the source document were saved as
   private notes"); better, make it a choice.
2. **`Browse.html` is a plaintext dump of every active clause, left permanently in the library
   folder.** The entry XML is plaintext too, so it is not a new *class* of exposure, but a single
   `.html` file is far more portable and will be **content-indexed by Windows Search**, making
   clause wording discoverable through the Start menu. Nothing ever deletes it, and it goes stale
   silently. See §11 — I think this feature should go.
3. **Path-based, not encrypted.** Correctly and explicitly stated in `READ_ME.txt`
   ("Offline operation is not encryption"). For a personal tool on a BitLocker'd laptop this is
   the right call — noted here only so the decision is visible rather than assumed.
4. **No complete data removal** (§7).

No injection, traversal or privilege issue was found. `Shell "explorer.exe " & Chr$(34) & path &
Chr$(34)` is safe — the path comes from a folder picker and Windows paths cannot contain `"`.

---

## 10. Test and validation evidence

`VALIDATION.md` claims 65 native Word assertions, 7 JavaScript assertions, and manual
setup/removal/reinstall lifecycle testing. Assessing the evidence itself:

**What holds up**

- The SHA-256 test vectors in `CLNativeTests` (empty string and `"abc"`) are **correct** — I
  verified both against real SHA-256. Anchoring the CryptoAPI wrapper to published vectors is
  exactly right and catches the BOM-offset class of bug.
- The tests exercise genuinely valuable things that only Word can answer: bold/italic survival,
  automatic and two-level outline numbering, whole tables, source immutability under tracked
  changes, comment exclusion, single-Undo restoration, protected-destination refusal,
  partial-table refusal, corruption isolation, and a backup→restore byte-identity check.
- The corruption test (`CLMoreTests:77-80`) is well constructed: it tampers with a saved file and
  asserts the healthy entry still lists while the bad one is reported.

**Evidence gaps**

- **No test results ship.** `native-results.txt`, `safety-results.txt` and `workflow-results.txt`
  are generated next to the test document and are explicitly excluded from the package. So the
  recipient cannot verify the "65 passed" claim; it rests entirely on the author's assertion.
  `BUILD.md`'s reason for excluding them is sound, but the consequence should be acknowledged.
- **Each suite aborts on the first unhandled error** (`On Error GoTo Fatal` → `GoTo Finish`), and
  the reported total is simply whatever ran. A run that died at assertion 5 reports "4 passed,
  1 failed" — it does not report that 60 assertions never executed.
- **The concurrency test is same-process.** `CLMoreTests:31-38` opens the lock with a second
  `FreeFile` in the *same* Word instance. VB's `Lock Read Write` does map to an exclusive share
  mode that blocks other processes, so the behaviour is probably right — but the cross-Word-
  process case that `PRODUCT_DECISIONS.md` claims is **not what was tested**.
- **The 7 JavaScript assertions are not in the source bundle**, so the catalogue's test harness
  cannot be inspected. (I re-derived and ran equivalent tests myself — see §6, §9.)
- **Nothing covers** the control-character class (§3.1), backup with a corrupt entry (§3.4),
  insertion into a large document (§4.2), a library with more than a handful of entries (§3.3),
  `CLBrowse` output, or the placeholder dialog.
- `CLWorkflowTests:49` mutates `ListGalleries(wdOutlineNumberGallery)` — a **global Word setting**
  — which is why `BUILD.md`'s "disposable Word session only" instruction is load-bearing and
  should be stated more emphatically.

---

## 11. Complexity, redundancy and unnecessary mechanisms

For a **single-user, single-machine, offline** tool, the save path carries five overlapping
integrity mechanisms:

| mechanism | what it uniquely protects against |
|---|---|
| `writer.lock` | two Word processes writing at once — **genuinely needed** |
| expected-revision check | two windows editing the same entry — **genuinely needed** |
| atomic write (`MoveFileEx`) | a partial file after a crash — **genuinely needed** |
| read-back verification (twice: in `CLWrite`, again in `CLAtomicWrite`) | a lying filesystem — **mostly redundant with the above** |
| `digest` | a hand-edited file — **net negative** (§3.6) |

Two or three would do. Similarly, the product contains **two hashing mechanisms**: `CLFileSha`
(hashes real bytes, used for the template copy — the robust one) and `CLDigest` (hashes an MSXML
re-serialisation, used for the data that matters — the fragile one).

Specific candidates for removal:

1. **The whole-document insertion rehearsal** (§4.2) — highest cost, no verified benefit.
2. **The per-entry `digest`** (§3.6) — or demote it to advisory with a repair path. XML
   well-formedness plus the atomic write already cover the realistic failure modes.
3. **`Browse.html` / the HTML catalogue.** It duplicates the library form's search with a *worse*
   feature set (no metadata, no private notes, no favourites, no filters), it goes stale silently,
   it adds a plaintext confidentiality surface that Windows Search will index, and the one thing
   it uniquely offers — getting wording out as plain text — is already served by "Insert plain
   text". Roughly 25 lines of generator plus a shipped promise to maintain. **Remove it, or
   replace it with a real export (§12).**
4. **`CLCaptureSelection`'s comment branch** duplicates `CLCaptureComment` (with slightly
   different validation — one checks for empty text, the other doesn't). One path.
5. **The three-stage code-generation build chain** (§8).
6. **Dead fields** `source`, `created`, `audience` — either surface them (they are useful, §5.3)
   or stop writing them.

---

## 12. Important missing capabilities

In rough order of value to the actual legal workflow:

1. **Quick insert from the ribbon** — a favourites gallery or drop-down. The single most
   frequent action currently takes the most clicks (§5.1).
2. **Better information at the point of reuse** — date, source document, topic/tags in the list;
   a formatted (or at least numbered) preview; a sortable date column (§5.2, §5.3).
3. **Comparison between revisions.** The product stores every revision forever and offers no way
   to see what changed. *"How does my current indemnity differ from the one I used in March"* is
   arguably the reason history exists. Word's own `Application.CompareDocuments` makes this nearly
   free, and it would be a genuinely distinctive capability.
4. **Duplicate detection at capture** (§5.5).
5. **A usage signal** — "last used", "used N times". The best relevance signal a precedent
   library can have, free to collect, and it would give the list a meaningful default sort.
6. **A repair path for corrupt entries** (§3.5).
7. **Full export** — every entry to a folder of `.docx` files plus a metadata CSV. This is the
   answer to "can I get my content out", and it is also the honest migration path if the product
   is ever retired (§7).
8. **Backup rotation / pruning**, and history capping (§3.2, §7).
9. **Complete uninstall**, including the support folder and registry value (§7).
10. **A product version**, surfaced in the Quick guide and recorded in `library.xml` (§8).
11. **Visible identification of the connected library** in the form — folder path and last-backup
    date in the status line (§7).
12. **A visually distinct editing draft** (§6).

---

## 13. Better alternative approaches

I am not proposing a rewrite. Two changes would remove most of the structural problems:

### 13.1 Store the payload beside the record, not inside it

```
entries/<guid>.xml     metadata only, 1–2 KB          ← listing and search read only this
entries/<guid>.docx    the clause as a real Word file ← opened only to preview or insert
history/<guid>-r7.xml  + <guid>-r7.docx
```

- Listing and searching a 1,000-entry library reads ~2 MB instead of ~78 MB.
- The library becomes directly inspectable: double-click a `.docx` and see the clause.
- Insertion can use `Range.InsertFile`, Word's oldest and best-tested insertion path, instead of
  `InsertXML` of a Flat OPC round trip — which also removes the exact-text gate (§4.3) and most
  of the style-import problem (§4.5), because Word's file-insert has mature style-conflict
  handling.
- `CLFileSha` (already written) verifies payloads by real bytes; `CLDigest` can go.
- Backup becomes a folder copy whose cost is proportional to real content.

This is a migration, not a rewrite: `CCLCaptureRich`'s cleaning pipeline is unchanged; it just
saves the cleaned temp document rather than serialising it.

### 13.2 Seriously evaluate Word Building Blocks before building more

Word already ships a precedent store: **Building Blocks** in a `.dotx`/`.dotm` gallery. They give
you native rich storage, native insertion, a gallery UI on the ribbon, survival across restarts,
and a single portable file — for free, maintained by Microsoft.

What they *don't* give you is full-text search, metadata (topic/tags/notes/family), history, and
a decent picker. Those are exactly what this product adds well.

So the question worth asking before the next increment is: **could this be a thin metadata-and-
search layer over Building Blocks, rather than a parallel storage engine?** That would delete
`CLStore`, `CLPlatform`'s file and crypto primitives, `CLAtomicWrite`, the lock, the digest, the
per-entry files and most of `CLRecovery` — and backup would become "copy one file".

The real trade-off is honest and worth weighing: a single container file is a single point of
failure and is harder to inspect, diff or partially recover than per-entry files. That is a
genuine argument **for** the current shape. But it should be a decision made deliberately, and I
can't see evidence in `PRODUCT.md` or `PRODUCT_DECISIONS.md` that the Building Blocks option was
evaluated and rejected — and a tool that duplicates a native Word capability needs to know why.

### 13.3 What I would *not* change

Do **not** move to SQLite, a JSON index, a local web app, a browser extension, a sync service, or
a tagging ontology. The current instinct — plain files, no dependencies, Word as the only writer,
no network — is correct and should be defended against feature pressure.

---

## 14. What is already well designed and should be preserved

- **The source document is never modified.** Non-negotiable for this audience, implemented
  correctly, and tested.
- **The disposable-copy capture pipeline.** Resolving tracked changes, comments, fields and
  hidden text in a throwaway document is exactly the right shape.
- **`w:sectPr` stripping.** Subtle, important, easy to have missed.
- **Explicit refusal over silent flattening** for embedded objects, floating shapes and partial
  tables.
- **The payload blocklist validator**, run both after cleaning and before insertion.
- **Private notes structurally separated from insertable content**, enforced in code.
- **"Not organised" as a reminder that never gates reuse.** The right philosophy, and rare.
- **Archive and recoverable Trash as visible user concepts**, with no hard delete anywhere.
- **The library identity marker** and the refusal to create a silent empty replacement.
- **"Marker written last"** in backup and restore.
- **Restore-as-a-new-revision** rather than rewriting history.
- **Byte-verified installation with rollback**, and a persistent support copy so removal remains
  possible.
- **`LeaveCurrent()` called on every navigation path** — no route out of an edited details panel
  loses work, and I could not find a gap.
- **The document-visible comment confirmation** that shows the text and says who will see it.
- **Honest documentation.** `VALIDATION.md`'s "Practical limits" section states what was not
  tested. That is rarer and more valuable than the tests themselves.
- **Verifiable release hygiene** — checksummed manifest, scrubbed metadata, shipped source that
  actually matches the binaries.

---

## 15. Release-readiness assessment

**Not ready for reliance on a professionally valuable repository. Suitable for continued personal
use on a small library with an independent backup, while the items in §16 are addressed.**

| Dimension | Assessment |
|---|---|
| Does it corrupt Word documents? | No evidence of it. Capture is read-only; insertion is guarded and rolled back. **Good.** |
| Does it lose data silently? | No. Failures are loud and the prior state is preserved. **Good.** |
| Does it *refuse* to accept data it should accept? | **Yes — §3.1.** Realistic content fails opaquely. |
| Does it recover from its own failures? | **Partly.** Backup aborts on one bad file (§3.4); a corrupt entry has no in-product repair (§3.5). |
| Does it scale to a multi-year repository? | **No, not as built** (§3.2, §3.3). |
| Is it pleasant after the novelty wears off? | **Probably not** (§5). |
| Is it honest about itself? | **Yes** — unusually so. |
| Is it maintainable in two years? | **Not as built** (§8). |

The three release blockers are §3.1 (capture fails on ordinary content), §3.4 (backup fails when
needed most) and §3.2/§3.3 (the library becomes unusable as it becomes valuable). Everything else
is improvement, not blocking.

---

## 16. Prioritised recommendations

### Must address before serious reliance

1. **Sanitise control characters in `CCLPlainWording`.** Strip or substitute everything outside
   the XML 1.0 white-list. *(§3.1 — ~6 lines, removes an entire class of opaque failure.)*
2. **Make backup and restore resilient to damaged files.** Copy what is healthy, copy damaged
   files verbatim without validating, report both counts. Never abort the whole operation.
   *(§3.4)*
3. **Stop reading the rich payload during listing and search.** Either move the payload to a
   sibling `.docx` (§13.1, preferred) or, as an interim, cache the digest and skip re-hashing
   unmodified files, and stop calling `RefreshEntries` after every metadata save. *(§3.2, §3.3)*
4. **Remove the whole-document insertion rehearsal** (`CLRich.bas:131-135` and the `beforeXml`
   logic in the error handler). Keep the payload probe. *(§4.2)*
5. **Add a loop guard to the hidden-text deletion loop.** *(§4.6 — 2 lines, removes a hang risk.)*
6. **Give corrupt entries a repair path** — a "Problem entries…" button offering restore-from-
   history — or change the error message to advice the user can actually follow. *(§3.5)*
7. **Disclose the source-comment import.** Tell the user at capture time, ideally with a choice.
   *(§9.1)*

### Should improve next

8. **Surface the data you already store**: `created`, `updated`, `source`; add a date column, a
   sort control, and `source` to the search fields. *(§5.3 — highest value per line of code in
   this list.)*
9. **Fix titles** — derive from the nearest source heading, or prompt once with the auto-title
   pre-filled. *(§5.4)*
10. **Add quick insert from the ribbon** (favourites gallery or recent list). *(§5.1)*
11. **Make the exact-text gate a warning, not a block**, with an "Insert anyway" option. *(§4.3)*
12. **Strip quotes from search queries**, and offer help on zero results. *(§6)*
13. **Remove the placeholder dialog**; jump the cursor to the first placeholder after insertion
    instead. *(§5.6)*
14. **Replace the two raw VBA errors** ("Permission denied", "The requested member of the
    collection does not exist") with written messages, and disable "Save wording" when the active
    document is not a draft. *(§6)*
15. **Improve the reuse preview** — show list numbering, and make the preview large enough to
    read a clause. *(§5.2)*
16. **Add a product version** to the code, `library.xml` and the Quick guide. *(§8)*
17. **Show the connected library folder** (and last backup date) in the form. *(§7)*
18. **Document the cross-reference freezing** in `READ_ME.txt`, and ideally notice it at capture.
    *(§4.4)*
19. **Abandon the code-generation build chain**; maintain the `.dotm` directly and export `.bas`
    files for version control. Remove the source-mutating "validate" step. *(§8)*

### Useful only if real usage justifies it

20. Revision comparison via Word's `CompareDocuments`. *(§12.3 — high value, but only once the
    library is large enough to have interesting history.)*
21. Duplicate detection at capture. *(§5.5 — matters from roughly 100 entries onward.)*
22. Usage counters and a "recently used" sort. *(§12.5)*
23. Full export to `.docx` + CSV. *(§12.7 — do this before ever sharing the tool with anyone else.)*
24. Backup rotation and history capping. *(§7 — defer until the growth is actually felt.)*
25. Complete uninstall. *(§7)*
26. Insertion into table cells. *(§4.7 — gather evidence on how often it is actually wanted.)*

### Do not change / already good

27. Source-document immutability and the disposable-copy capture pipeline.
28. `w:sectPr` stripping and the payload blocklist validator.
29. Explicit refusal over silent flattening.
30. Private notes structurally separated from insertable content.
31. "Not organised" as a non-gating reminder; Archive and recoverable Trash; no hard delete.
32. The library identity marker and the refusal to create a silent empty replacement.
33. "Marker written last" in backup and restore; restore-as-a-new-revision.
34. The writer lock and the expected-revision check.
35. `LeaveCurrent()` on every navigation path.
36. Offline, no-network, no-account, no-runtime-dependency operation.
37. The honesty of `VALIDATION.md` and the unsigned/limits disclosures.
38. SHA test vectors anchored to published references.

### Recommendations that are *not* code changes

- **§3.2 should be confirmed before it is engineered.** One `?Len(Selection.WordOpenXML)` in the
  Immediate window decides how urgent recommendation 3 is.
- **§13.2 (Building Blocks) is a question to answer, not a change to make.** Write down why the
  custom store is preferred, or discover that it isn't.
- **§5 is a usage question.** Keep a tally for a month of how often the library is opened, how
  often an item is inserted, and how many entries still carry an auto-title. If reuse is rare,
  the answer is recommendation 10, not more storage engineering.
- **Recommendations 24, 25, 26 should be deferred** until real usage shows they matter.

---

## 17. What I would do differently from first principles

Starting today, for one lawyer, offline, on Windows, in Word, I would hold onto exactly three
properties of the current product — **never touch the user's document**, **capture in one click**,
and **no dependency beyond Word** — and change nearly everything else.

**The job is not "store clauses". It is "at the moment of drafting, find wording I trust and put
it in correctly."** Every design decision should be judged against the ten seconds in which a
lawyer is mid-sentence and wants their indemnity clause. Measured that way, the current product
spends most of its engineering on the storage layer and almost none on that ten seconds.

Concretely:

- **Storage:** metadata as small files; the clause as a real `.docx` beside it. Small, portable,
  inspectable, greppable, and readable without the tool. Insert with `Range.InsertFile`.
  Backup is a folder copy. No digest, no Flat OPC escaping, no atomic-write ceremony beyond a
  temp-file rename. Keep the writer lock and the expected-revision check — they earn their place.
- **History:** keep the previous payload file and cap it at the last ten revisions. Compare two
  revisions with Word's own Compare. Full immutable history of a 76 KB envelope buys almost
  nothing over that, and costs gigabytes.
- **Capture:** one click, but with one field — a name, pre-filled from the nearest heading,
  dismissible with Enter. That single second of friction is what separates a searchable library
  from a pile of truncated first lines. Show a real confirmation, not a status-bar message. Warn
  on a near-duplicate. Ask before absorbing the source document's comments.
- **Reuse:** the primary interface is a **ribbon gallery of favourites and recents**, one click,
  with a proper preview. The search window is the secondary interface for when the gallery isn't
  enough. Today those priorities are inverted.
- **Retrieval:** rank results instead of filtering them — title match above tag match above body
  match, recently-used first. Match on word boundaries to kill the `ip`/`Recipient` noise. Handle
  quoted phrases. Show date, source and topic in the list.
- **Trust:** show what is connected and when it was last backed up, in the window, always. Make
  export a first-class feature from day one — a lawyer will not invest years in a repository they
  cannot get out of.
- **Engineering effort:** roughly the inverse of the current allocation. Less on proving each
  save is atomic five different ways; much more on the six lines of preview and the one-click
  insert that determine whether the tool is opened at all.

And one honest test I would apply before building anything further: **spend a week using Word
Building Blocks for the same job.** If that feels almost adequate, the right product is a
metadata-and-search layer over them, and most of `CLStore`, `CLPlatform` and `CLRecovery` should
never be written.

---

## 18. Proposed next direction

Not a feature list — the smallest coherent set of actions that makes this safer, simpler and more
likely to survive.

**Step 0 — measure, before building (one hour).**
Run `?Len(Selection.WordOpenXML)` on one short clause. Count how many entries are in the library,
how many still carry an auto-generated title, and how many times a clause has actually been
inserted this month. These four numbers determine whether Step 2 is urgent and whether Step 3 is
the real problem. **Do not skip this.**

**Step 1 — stop the three ways the product can fail you (one focused change).**
Sanitise control characters (§3.1). Make backup skip-and-report instead of abort (§3.4). Add the
loop guard (§4.6). Give corrupt entries a repair route (§3.5). These are small, independent, and
together they remove every finding in this audit that can cost the user work.

**Step 2 — take weight out (one focused change).**
Delete the whole-document rehearsal (§4.2). Move the rich payload to a sibling `.docx` and switch
insertion to `Range.InsertFile` (§13.1). Retire the per-entry digest in favour of `CLFileSha` on
the payload (§3.6). Remove `Browse.html` (§11.3) — or, better, replace it with a real export,
which serves the same need and more. Stop calling `RefreshEntries` after every metadata save.

This step makes the product *smaller*. It removes roughly a third of the storage machinery, the
scale ceiling, the exact-text gate, the style-import problem and one confidentiality surface at
the same time.

**Step 3 — make it worth opening (one focused change).**
Surface `created`, `updated` and `source` with a sortable date column (§5.3). Fix auto-titles
(§5.4). Add a favourites gallery to the ribbon (§5.1). Delete the placeholder dialog (§5.6). Fix
quoted search (§6). These are shallow changes with the largest effect on whether the repository
is still being maintained in three years.

**Step 4 — decide, don't drift.**
Write down the answer to §13.2: is this a storage engine, or a search-and-metadata layer over
Word's own? And add a version number before there is a second version to distinguish.

**Then stop and use it for a quarter.** The honest finding of this audit is that the product's
correctness is well ahead of its usefulness, and further engineering on integrity will not change
whether a lawyer keeps opening it. The next genuinely useful input is not another mechanism — it
is a quarter's worth of real usage data.
