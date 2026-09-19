# Clause Library 2.0 — what to do next

## The honest position first

I could not run Word. This machine is Linux; there is no Word, no VBA, no
Windows. So:

- Every line of the product has been written, reviewed, and put through a
  static checker that stands in for the VBA compiler (block structure, calls
  resolving, ribbon callbacks, form controls, project boundaries). That checker
  is itself mutation-tested: fourteen deliberate breakages, all caught.
- The search ranking and the exported page's JavaScript were **executed** here
  and pass.
- **Nothing has been run inside Word.** Capture, insertion, numbering, tables,
  Track Changes, undo, installation — none of that has been observed working.

That is why the product now contains a **self-check**: about eighty assertions
that exercise the storage layer and Word itself against a throwaway library.
It is the thing that turns "should work" into "does work", and it takes about
a minute.

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

## Things I decided without asking you

- **Kept it as Word VBA.** An Office Add-in would be the modern answer but needs
  a server and a manifest, which kills offline working. For this job VBA is
  still right.
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
