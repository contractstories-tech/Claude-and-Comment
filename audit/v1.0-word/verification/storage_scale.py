import zipfile, hashlib, io, os, time
from xml.sax.saxutils import escape
Z='/tmp/claude-0/-home-user-Claude-and-Comment/811b643d-129e-5b6d-a70e-f031dc43493d/scratchpad/audit/pkg/ClauseLibraryPersonal.dotm'
z=zipfile.ZipFile(Z)
# Range.WordOpenXML returns a Flat OPC package: the parts Word always emits with a range.
parts=['word/document.xml','word/styles.xml','word/theme/theme1.xml','word/settings.xml',
       'word/webSettings.xml','word/fontTable.xml']
sizes={p:len(z.read(p)) for p in parts}
base=sum(sizes.values())
print("Parts Word emits in a Flat OPC payload (sizes taken from the shipped template itself):")
for p,s in sorted(sizes.items(), key=lambda kv:-kv[1]):
    print(f"   {p:28} {s:7,d} bytes")
print(f"   {'numbering.xml (absent here; a numbered clause adds it)':28} ~{4000:7,d} bytes est.")
print(f"   {'--- fixed overhead per clause':28} {base:7,d} bytes\n")

clause = ("The Supplier shall indemnify and hold harmless the Customer against all "
          "Losses arising out of or in connection with any breach of this Agreement.")
print(f"Actual clause wording                       {len(clause):7,d} bytes")
flat = base + 4000 + len(clause)*4          # +4000 numbering, x4 for WordprocessingML run markup
print(f"Flat OPC payload (conservative)             {flat:7,d} bytes")
stored = len(escape(flat*'x'))              # placeholder; do it properly below
sample = "<w:p><w:r><w:t>"+clause+"</w:t></w:r></w:p>"
inflate = len(escape(sample))/len(sample)
print(f"XML-escaping inflation when nested as text  x{inflate:.2f}")
entry = int(flat*inflate)
print(f"Stored size of ONE entry file               {entry:7,d} bytes  ({entry/1024:.0f} KB)\n")

print("Library growth, with 8 edits per clause over its life (history keeps a full copy each time):")
print(f"{'clauses':>8} {'entries MB':>11} {'history MB':>11} {'library MB':>11} {'read+hashed per refresh':>24}")
for n in (100,300,1000,3000):
    ent=n*entry; hist=n*8*entry
    print(f"{n:>8} {ent/1e6:>11.1f} {hist/1e6:>11.1f} {(ent+hist)/1e6:>11.1f} {ent/1e6:>21.1f} MB")

print("\nWork CLList() performs on EVERY library refresh (open form, capture, save details,")
print("archive, trash, favourite tick, wording save):  read + UTF-8 decode + DOM parse")
print("+ re-serialise /entry/data + UTF-8 re-encode + SHA-256, for every entry file.\n")
for n in (100,300,1000):
    payload=b'x'*entry
    t0=time.perf_counter()
    for _ in range(n): hashlib.sha256(payload).hexdigest()
    t=time.perf_counter()-t0
    print(f"  {n:>5} entries: {n*entry/1e6:6.1f} MB hashed. Raw SHA-256 alone in C: {t*1000:7.1f} ms")
print("\n  (VBA does this through ADODB.Stream + MSXML + CryptoAPI per entry; the DOM parse and")
print("   re-serialisation dominate and are orders of magnitude slower than the raw hash above.)")
