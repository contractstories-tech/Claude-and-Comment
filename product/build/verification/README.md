# Checks that run without Word

Everything the product does inside Word needs Word. These are the parts that
can be proved before that, and they run on any machine with Python 3 and Node.

| what | why it matters |
|---|---|
| `../check-sources.py` | Stands in for the VBA compiler: block structure, calls resolving, ribbon callbacks, form controls, per-project module sets. Mutation-tested — deliberately breaking the source in fourteen different ways is caught every time. |
| `search_model.py` | A faithful port of the new ranked search. Proves quoted phrases work (v1 returned nothing for `"force majeure"`) and that a short word like `ip` ranks the intellectual property clause above `Principal` and `Recipient`. |
| `export_page.js` | Extracts the exact JavaScript `CLExport.bas` writes into the exported index and runs it in a real DOM. |

Run them:

```
python3 ../check-sources.py
python3 search_model.py
node export_page.js          # needs: npm install jsdom
```

The self-check that ships inside the product (`word/CLSelfCheck.bas`, about
eighty assertions) covers everything these cannot: Word's own capture,
formatting, numbering, tables, tracked changes, undo, protection and the whole
storage layer against a throwaway library. That one has to be run in Word.
