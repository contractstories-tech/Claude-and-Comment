from lxml import etree
# Characters Word's Range.Text can yield, and what CCLPlainWording does with them
cases = {
 1:  "inline picture / drawing anchor",
 2:  "footnote or endnote reference mark",
 5:  "comment (annotation) reference mark",
 7:  "table cell / row end  -> converted to TAB by CCLPlainWording",
 9:  "tab",
 11: "manual line break     -> converted to LF by CCLPlainWording",
 12: "page break / section break",
 13: "paragraph mark        -> converted to LF by CCLPlainWording",
 14: "column break",
 19: "field begin",
 21: "field end",
 30: "non-breaking hyphen (Ctrl+Shift+-)",
 31: "optional (soft) hyphen",
}
print(f"{'chr':>4}  {'raw in XML':>11}  {'&#n; ref':>9}   meaning")
print("-"*86)
for n, what in sorted(cases.items()):
    ch = chr(n)
    raw = ref = "ok"
    for form, label in ((ch, 'raw'), (f"&#{n};", 'ref')):
        doc = f"<entry><data><plain>A{form}B</plain></data></entry>".encode()
        try:
            etree.fromstring(doc)
            r = "ACCEPT"
        except Exception:
            r = "REJECT"
        if label == 'raw': raw = r
        else: ref = r
    print(f"{n:>4}  {raw:>11}  {ref:>9}   {what}")
