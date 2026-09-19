# Clause Library 2.1 — what to do next

## The honest position first

I could not run Word. This machine is Linux; there is no Word, no VBA, no
Windows. So:

- Every line of the product has been written, reviewed, and put through a
  static checker that stands in for the VBA compiler: block structure, calls
  resolving, **argument counts**, ribbon callbacks, form controls, class
  imports, project boundaries and the no-egress rule. That checker is itself
  mutation-tested — thirty deliberate breakages across those categories, all
  caught.
- The search ranking and the exported page's JavaScript were **executed** here
  and pass.
- **Nothing has been run inside Word.** Capture, insertion, numbering, tables,
  Track Changes, undo, installation — none of that has been observed working.

That is why the product now contains a **self-check**: about eighty assertions
that exercise the storage layer and Word itself against a throwaway library.
It is the thing that turns "should work" into "does work", and it takes about
a minute.

## What the four reviews changed (version 2.1)

Three genuinely independent reviews of **1.0** were read in full, plus a fourth
zip that turned out to be this product's own 2.0 handed back. Most of what they
found, 2.0 had already fixed. Seven things had not been, and three of those were
serious enough to be release-blocking:

- **What you read was not necessarily what got inserted.** The window kept the
  record it loaded; Insert re-read from disk. Save an edit, restore an older
  version, or let another Word window write, and the preview and the insertion
  could disagree. Every action is now bound to the version you actually looked
  at. This was the single most important defect in the product.
- **Ctrl+S in a wording draft wrote a file and left the library untouched**,
  while looking exactly like saving your clause. Save now means save to the
  library, and closing asks save / discard / cancel.
- **Plain-text insertion skipped the destination checks** the formatted path
  makes. One preflight now serves both.
- Previous-version snapshots were written non-atomically; the library's
  identity file had no spare copy and no way back; the folder-length limit was
  two numbers that nearly disagreed; and a busy library failed instantly
  instead of waiting a moment.

Also added, because the reviews were right that they were missing: permanent
delete (the only irreversible action in the product), updating without
uninstalling, a flag when an item's details no longer describe its wording,
and a visible backup age.

Three recommendations were **declined**, with reasons in
`docs/What changed in 2.1.txt`: banning images, removing the setup copy that
makes uninstallation possible, and adding a cached metadata index.

## The three steps

**1. Build it** (once, on your Windows machine, ~5 minutes)

Full instructions in `product/BUILD.md`. In short: tick *Trust access to the
VBA project object model* in Word's Trust Center, open a blank document, press
Alt+F11, import `product/build/CLBuild.bas`, put the cursor in
`BuildClauseLibrary`, press F5, then untick the Trust Center setting.

Word writes the macro project itself. That is deliberate — it is the only way
to be certain the binary is right, and it means **nobody you send the result to
needs that setting, or any other special setting.**

**2. Check it** (~2 minutes)

Open `product/dist/Start Here.docm`, set it up, then on the Clause Library tab
choose **Options → Run self-check**.

- All checks pass → the folder `product/dist` is what you circulate.
- Anything fails → send me the log. It names exactly what failed and where.

**3. Live with it for a few weeks before sending it widely**

Capture thirty or forty real clauses. Use *Insert* rather than opening the
library. See whether the names you get are ones you recognise a month later.
That is the thing no amount of engineering settles.

## What actually changed, in one paragraph each

**It could not store some ordinary wording.** Footnotes, inline images, page
breaks and non-breaking hyphens all produce characters XML cannot hold, and
capture failed with a message about library files. Fixed, and the self-check
proves it end to end with a real footnote.

**Backup refused to run when a file was damaged** — exactly when you need it.
Backup and restore now copy what they can and report the rest.

**A damaged item was unreachable.** The error said to use its history, but the
item was hidden from the list, so its history could not be opened. There is now
a *Needs attention* view and a repair.

**It would have buckled as it grew.** Every item cost 60–100 KB whether it was
one sentence or ten pages, and all of it was re-read and re-checksummed on
every refresh. The Word formatting now lives in its own file beside the record:
a thousand clauses is about 2 MB to search rather than 80 MB. Existing
libraries convert themselves, after taking a backup.

**Every insertion copied your entire contract** into a hidden document to
rehearse — and never checked the result. Gone. The cheap, useful check kept.

**It was not worth opening.** There is now one-click Insert and Favourites on
the ribbon; the list shows topic, type and when you last used it and can be
sorted; names are suggested from the heading the wording sits under instead of
being the first 90 characters; the preview shows the numbering Word leaves out
of plain text; quoted phrases work in search and results are ranked; duplicates
are noticed; versions can be compared with Word's own comparison; and there is
a real Export that writes everything out as ordinary Word files.

**Two things were removed.** The `[placeholder]` warning, which fired on nearly
every real clause and taught people to dismiss the warnings that matter — the
first placeholder is now selected for you after insertion instead. And the
generated `Browse.html`, which duplicated the library window with fewer
features and left a plain-text copy of every clause in the library folder for
Windows Search to index. Export does that job properly.

**What was already right was left alone**: your document is never modified when
you capture; tracked changes resolve in a throwaway copy; page setup is
stripped so a clause cannot alter a contract's layout; unsafe content is
refused out loud rather than flattened; private notes are separated in code,
not by convention; nothing is ever deleted outright; the backup completion
marker is written last; restoring a version adds one rather than rewriting
history.

## On "offline" meaning no data egress

You clarified that offline means the tool must not send data outside, because
of data-restriction policy and confidential agreement text — and that running
on corporate server infrastructure is fine. That is a better-defined
requirement than the one I designed against, and the build already satisfies
it. I have now made it provable rather than asserted:

- **The build refuses to produce a package** if the source contains any HTTP
  client, email API, download command, hard-coded address, or code that
  re-enables external entity resolution in the XML parser. `check-sources.py`
  names the file and line. Ten deliberate attempts to smuggle network code in
  were each rejected.
- **Every path is checked before it reaches Word or Explorer.** Word opens a
  web address as readily as a file; anything with a URL scheme is now refused.
- **Sync folders are the one real route out**, and the tool cannot prevent
  them — OneDrive, SharePoint, Dropbox and the rest copy files out by design.
  It now detects them and asks before using one, naming the service and saying
  what will be copied. It does not block the choice, because in some firms the
  synced location is the approved one. The default library location is not
  synced by Known Folder Move.
- **`docs/For your IT department.txt`** is written to be handed straight to
  InfoSec: what it installs and where, the full list of COM objects it creates
  (five, all local), what it refuses to accept from captured content, and how
  to verify the no-egress property themselves in one command.

What your clarification **does** change, for later: a shared team library over
a file share or an internal server is now on the table, where before I had
ruled it out. I have not built it, and I would not build it until you have used
the personal version for a while — shared precedent raises questions
(who may edit, whose approval a clause carries, what happens to someone's
private notes) that are governance questions, not engineering ones. The
storage design does not block it: plain per-entry files on a UNC share with the
existing writer lock would carry a small team as-is.

What it does **not** change: Word VBA is still the right vehicle. An Office
Add-in would need a web server and a manifest, and buys nothing here.

## Things I decided without asking you

- **Kept it as Word VBA.** An Office Add-in would be the modern answer, but it
  needs a web server and a manifest and gains nothing for a personal library.
  Now that server-side deployment is acceptable to you, it stays worth
  revisiting only if a shared team library is what you actually want.
- **Kept per-entry files rather than moving to Word Building Blocks.** Building
  Blocks would delete most of the storage code, but a single container file is
  a single point of failure and cannot be inspected, diffed or partly
  recovered. For something you are meant to trust for years, I'd rather have
  plain files. Worth revisiting once you know how you actually use it.
- **Kept the digest, but made it repairable.** In 1.0 a record edited by hand
  was permanently unreadable. Now it is flagged, still shown, and can be
  accepted or restored.
- **Capture now asks for a name.** It costs one keystroke and it is the
  difference between a library you can search and a list of near-identical
  opening lines. It can be turned off in the settings if you disagree.

## What I would still not claim

The self-check covers a great deal but not everything: unusual numbering
schemes, conflicting custom styles, every Office build, managed corporate
environments, and how this behaves on a library of several thousand items.
The design is now built for that scale; it has not been observed at it.
