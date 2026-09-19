# Faithful port of CLStore.CLMatches (VBA) -- same algorithm as the Browse.html JS
FIELDS = ["title","plain","topic","tags","family","role","applicability","notes","comment"]
def CLMatches(e, query, view):
    s = e.get("state","")
    if view == "Archived":
        if s != "Archived": return False
    elif view == "Trash":
        if s != "Trash": return False
    else:
        if s != "Active": return False
    if view == "Not organised" and e.get("organised") == "1": return False
    if view == "Favourites" and e.get("favourite") != "1": return False
    text = ""
    for k in FIELDS: text += " " + e.get(k,"")
    text = text.lower()
    for token in query.strip().lower().split(" "):      # VBA: Split(..., " ")
        if len(token) > 0 and token not in text: return False
    return True

lib = [
 {"state":"Active","title":"Intellectual property licence","plain":"The Licensor grants the Licensee a non-exclusive licence to the Background IP."},
 {"state":"Active","title":"Principal contractor","plain":"The Principal shall procure that each participant complies."},
 {"state":"Active","title":"Recipient obligations","plain":"The Recipient shall keep Confidential Information secret."},
 {"state":"Active","title":"Term and termination","plain":"This Agreement continues for the Initial Term unless determined earlier."},
 {"state":"Active","title":"Governing law","plain":"This Agreement is governed by the laws of England and Wales."},
 {"state":"Active","title":"Force majeure","plain":"Neither party is liable for delay caused by a Force Majeure Event."},
]
def hits(q, view="All clauses and comments"):
    return [e["title"] for e in lib if CLMatches(e,q,view)]

print("PRIMARY WORD SEARCH (CLStore.CLMatches) -- substring AND, no word boundaries\n")
for q in ["ip","term","law","act","force majeure",'"force majeure"',"Force  Majeure","  ip  "]:
    h = hits(q)
    print(f"  query {q!r:20} -> {len(h)} hit(s): {h}")

print("\nEmpty / whitespace query:")
for q in ["", "   "]:
    print(f"  query {q!r:8} -> {len(hits(q))} hit(s)  (whole active library)")

print("\nView filter cross-check:")
lib2=[{"state":"Active","organised":"0","favourite":"0","title":"new capture","plain":"x"},
      {"state":"Active","organised":"1","favourite":"1","title":"tidy fave","plain":"x"},
      {"state":"Archived","organised":"1","favourite":"1","title":"archived fave","plain":"x"}]
for view in ["All clauses and comments","Not organised","Favourites","Archived","Trash"]:
    got=[e["title"] for e in lib2 if CLMatches(e,"",view)]
    print(f"  view {view!r:28} -> {got}")
