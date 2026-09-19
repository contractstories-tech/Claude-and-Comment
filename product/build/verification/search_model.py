#!/usr/bin/env python3
"""Faithful port of the new search in CLStore.bas (CLNormalise, CLTerms,
CLFieldScore, CLScore), run against queries a lawyer would actually type.

The VBA is the product; this is a transliteration used to check the algorithm
behaves before it ever reaches Word. Keep the two in step by hand."""
import re, sys

def CLNormalise(v):                                   # lower, punctuation -> space, space-wrapped
    return " " + re.sub(r"[^a-z0-9]+", " ", v.lower()).strip() + " "

def CLTerms(query):                                   # honours "quoted phrases"
    out, buf, in_quote = [], "", False
    for ch in query.strip():
        if ch in ('"', '“', '”'):
            if in_quote and buf.strip(): out.append(CLNormalise(buf))
            if in_quote: buf = ""
            in_quote = not in_quote
        elif ch == " " and not in_quote:
            if buf.strip(): out.append(CLNormalise(buf))
            buf = ""
        else:
            buf += ch
    if buf.strip(): out.append(CLNormalise(buf))
    return out

def CLFieldScore(hay, needle):
    if len(hay) <= 2: return 0
    if needle in hay: return 3                        # whole word or exact phrase
    if needle[:-1] in hay: return 2                   # start of a word
    if needle[1:-1] in hay: return 1                  # inside a word
    return 0

WEIGHTS = [("sTitle", 6), ("sMeta", 4), ("sGuide", 2), ("sBody", 1), ("sSource", 1)]

def CLScore(row, terms):
    if not terms: return 1
    total = 0
    for t in terms:
        best = max(CLFieldScore(row[f], t) * w for f, w in WEIGHTS)
        if best == 0: return 0                        # every term must appear somewhere
        total += best
    if row.get("favourite") == "1": total += 2
    return total

def row(title, plain, topic="", tags="", applicability="", notes="", source="", favourite="0"):
    return {"title": title, "sTitle": CLNormalise(title),
            "sMeta": CLNormalise(f"{topic} {tags} Clause"),
            "sGuide": CLNormalise(f"{applicability} {notes}"),
            "sBody": CLNormalise(plain), "sSource": CLNormalise(source),
            "favourite": favourite}

LIB = [
 row("Intellectual property licence", "The Licensor grants a non-exclusive licence to the Background IP.", "IP", "licence"),
 row("Principal contractor duties", "The Principal shall procure that each participant complies.", "Construction"),
 row("Recipient confidentiality", "The Recipient shall keep Confidential Information secret.", "Confidentiality"),
 row("Force majeure", "Neither party is liable for delay caused by a Force Majeure Event.", "Risk", "boilerplate"),
 row("Liability cap - preferred", "Total liability shall not exceed 100% of the Charges paid.", "Liability", "cap", "Our opening position", favourite="1"),
 row("Uncapped indemnity - resist", "The Supplier shall indemnify without limit.", "Liability", "indemnity", "Counterparty wording; push back"),
 row("Termination for convenience", "Either party may terminate on 30 days written notice.", "Term", "exit"),
]

def search(q):
    scored = [(CLScore(r, CLTerms(q)), r["title"]) for r in LIB]
    hits = sorted([s for s in scored if s[0] > 0], key=lambda x: -x[0])
    return hits

CASES = [
 ('"force majeure"',  lambda h: [t for _, t in h] == ["Force majeure"],           'a quoted phrase finds the clause (v1 returned nothing)'),
 ("force majeure",    lambda h: h[0][1] == "Force majeure",                        "the same words unquoted rank it first"),
 ("ip",               lambda h: h[0][1] == "Intellectual property licence",        'short word "ip": the right clause outranks Principal/Recipient'),
 ("ip",               lambda h: len(h) >= 3,                                       "...while the loose matches are still reachable, not hidden"),
 ("indemnity",        lambda h: h[0][1] == "Uncapped indemnity - resist",          "a word in the name beats the same word in the body"),
 ("liability cap",    lambda h: h[0][1] == "Liability cap - preferred",            "two words narrow, and the favourite ranks first"),
 ("LIABILITY",        lambda h: len(h) == 2 and h[0][1] == "Liability cap - preferred",  "capitals make no difference"),
 ("counterparty",     lambda h: [t for _, t in h] == ["Uncapped indemnity - resist"], "private notes are searched"),
 ("frustration",      lambda h: h == [],                                           "a word that appears nowhere finds nothing"),
 ("",                 lambda h: len(h) == len(LIB),                                "an empty search shows everything"),
 ("terminat",         lambda h: h[0][1] == "Termination for convenience",          "a part-word still finds it"),
 ('"liability cap"',  lambda h: h[0][1] == "Liability cap - preferred",            "a quoted phrase must appear as a phrase"),
 ('"cap liability"',  lambda h: h == [],                                           "...and the words in the wrong order do not match it"),
]

fails = 0
print(f"{'query':22} {'expectation':62} result")
print("-" * 118)
for q, ok, why in CASES:
    h = search(q)
    good = ok(h)
    fails += 0 if good else 1
    top = ", ".join(t for _, t in h[:3]) or "(nothing)"
    print(f"{('pass' if good else 'FAIL'):5}{q!r:18} {why:62} {top[:44]}")
print()
print("all search checks passed" if not fails else f"{fails} FAILED")
sys.exit(1 if fails else 0)
