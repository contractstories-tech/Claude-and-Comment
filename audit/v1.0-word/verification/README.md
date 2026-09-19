# Reproducible checks run during the Clause Library 1.0 audit

Everything here runs on Linux with Python 3 + `lxml`, and Node + `jsdom` for the one JS test.
None of it requires Word — that is the point: these are the findings that could be verified
without it. Every claim about Word's own behaviour in `../AUDIT.md` is labelled as reasoning,
not observation.

| script | what it establishes | audit section |
|---|---|---|
| `xml_control_chars.py` | XML 1.0 rejects the control characters `Range.Text` emits (footnote marks, inline pictures, page breaks, non-breaking/optional hyphens), raw **and** as numeric references. `CCLPlainWording` sanitises only three of them. | §3.1 |
| `search_semantics.py` | Faithful port of `CLStore.CLMatches`. Shows quoted phrases return zero hits, and that unbounded substring matching makes `ip` match *Principal* and *Recipient*, `act` match *contractor*. | §6 |
| `storage_scale.py` | Per-entry storage cost, using the shipped template's own part sizes as the measure of what a Flat OPC payload contains. Projects library and history growth. | §3.2, §3.3 |
| `browse_html_injection.py` | Re-implements `CLHtml` + `CLBrowse` string building and generates `Browse.html` from adversarial clause text. | §9 |
| `browse_html_search.js` | Runs the generated page's real JavaScript in a DOM (jsdom) against realistic legal queries. | §6, §9 |

Checks performed directly with shell tooling rather than scripts, also reproducible:

- `sha256sum -c MANIFEST_SHA256.txt` in the package folder — all five files verify.
- VBA extracted from both shipped binaries with `oletools` and diffed against `src/word/*.bas`:
  identical apart from VBA's own identifier-case normalisation.
- `docProps/core.xml` and `app.xml` in both binaries carry no author, company or
  `lastModifiedBy`; the author's Windows username appears nowhere in the shipped binaries
  (it does appear in `src/tools/make-builder.cjs`).
- Neither `vbaProject.bin` contains a digital-signature stream, and neither project is
  password-locked — matching what `VALIDATION.md` claims.
- Both packages are valid OOXML zips with `[Content_Types].xml` as the first entry.
- The SHA-256 vectors asserted in `CLNativeTests.bas` are correct.
