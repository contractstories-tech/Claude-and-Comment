# Building Clause Library

The finished product is two files: `ClauseLibraryPersonal.dotm` and
`Start Here.docm`. **Word writes the macro project itself** — that is the only
way to be certain the binary is correct, so the build has to run on Windows
with desktop Word once. After that, the two files it produces need nothing
special from anyone you give them to.

## What is where

```
product/
  word/            the product. Eight modules and two windows. This is the source of truth.
  build/
    CLBuild.bas         the builder you import into Word and run
    make-seeds.py       rebuilds the seed packages (ribbons, content types, the Start Here page)
    check-sources.py    static checks that stand in for a VBA compiler
    seed/               the two seed packages the builder fills in
    verification/       checks that run without Word
  docs/            the two text files that ship beside the product
```

`word/` is the only place product behaviour lives. `build/seed/*` carry the
ribbon XML and the page the user reads; they also carry a working macro project
purely so Word will open them, and the builder replaces every module in it.

## Before you build

```
python3 build/check-sources.py
```

Run this after any change to `word/`. There is no VBA compiler here, so it
stands in for one: block structure, every call resolving, every ribbon callback
present with the right shape, every control the forms touch being created, no
duplicate public names, no keyword used as a line label, and — the one that is
easy to get wrong — no module in the small setup project calling something only
the full runtime project carries. It exits non-zero if anything is wrong.

`build/verification/` holds checks that run without Word at all: the search
ranking, and the JavaScript in the exported index page.

## Building

1. In Word: **File → Options → Trust Center → Trust Center Settings → Macro
   Settings**, and tick **Trust access to the VBA project object model**.
   This is needed to *build*. Nobody you give the finished files to ever needs it.
2. Open a **new, empty** Word document. Press **Alt+F11**.
3. **File → Import File…** and choose `build/CLBuild.bas`.
4. Put the cursor inside `BuildClauseLibrary` and press **F5**.
5. Untick the Trust Center setting again.

It writes `product/dist/` containing the two built files, the two documents and
`MANIFEST_SHA256.txt`. It refuses to build if `check-sources.py`'s equivalent
checks fail inside Word, and the log is in `dist/build-log.txt`.

The checksums in the manifest come from Windows' own `certutil`, deliberately
not from the product's own checksum code, so the manifest is not vouched for by
the thing it is meant to check.

## Before you give it to anyone

1. Open `dist/Start Here.docm` and set it up.
2. On the Clause Library tab: **Options → Run self-check**.
3. Only pass it on if every check passes. The log lands in your temp folder and
   is the thing to send if it does not.

## Rebuilding the seeds

Only needed if the ribbon, the content types or the Start Here page change:

```
python3 build/make-seeds.py <folder containing a 1.0 package>
```

It takes the ribbon definitions from `make-seeds.py` itself, validates them,
strips document metadata, and checks the result is a sane Office package.

## Notes for whoever maintains this next

- **Do not hand-edit the built `.dotm`.** Change `word/`, run the checks, rebuild.
  The built files are outputs, not source.
- Module names matter. `CLBuild.bas` imports by name and the setup project takes
  only `CLPlatform`, `CLStore`, `CLSetup` plus a small shim. If you add a call
  from one of those three into another module, `check-sources.py` will tell you.
- Form layout is decided at run time in `LibraryForm.LayOut`, not at design time.
  The builder only creates the controls. To move something, edit the form source.
- Adding a control means adding it in **two** places: the `Add designer, ...`
  line in `CLBuild.bas` and wherever it is used in the form source.
  `check-sources.py` fails if they disagree.
- `CL_VERSION` in `CLPlatform.bas` is the one place the version is written.
- There are no Node, Python or network dependencies in the product. Python is
  used here only to build the seeds and run the checks; it is not needed to
  build the product in Word, and not needed to run it.
