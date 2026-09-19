#!/usr/bin/env python3
"""Static checks over the VBA source.

There is no VBA compiler here, so this stands in for one: block structure,
every call resolving to something that exists, every ribbon callback present
with the right shape, and every control the forms touch actually being created
by the builder. It is not a compiler, but it catches what one would shout about.
"""
import re, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORD, BUILD = ROOT / "word", ROOT / "build"

# Word / VBA / MSForms names we are entitled to call without defining.
BUILTIN = set("""
abs array asc ascw atn cbool cbyte ccur cdate cdbl cdec chr chrw cint clng csng cstr cvar cverr
date dateadd datediff datepart dateserial datevalue day dir doevents environ eof err error exp
filelen filter fix format freefile getsetting hex hour iif inputbox input instr instrrev int ipmt
isarray isdate isempty iserror ismissing isnull isnumeric isobject join lbound lcase left len lenb
loc lof log ltrim mid midb minute month monthname msgbox now oct partition qbcolor rgb right rnd
round rtrim savesetting second seek sgn shell sin space spc split sqr str strcomp strconv string
strptr strreverse switch tan time timer timeserial timevalue trim typename ubound ucase val
vartype weekday weekdayname year deletesetting createobject getobject cvdate asch loadpicture
setattr getattr fileattr curdir chdir mkdir rmdir kill name rset lset partition
""".split())

def modules():
    out = {}
    for p in sorted(list(WORD.glob("*.bas")) + list(WORD.glob("*.txt")) + list(WORD.glob("*.cls"))):
        out[p.name] = p.read_text(encoding="utf-8").replace("\r\n", "\n")
    return out

def logical_lines(text):
    """Join VBA line continuations, drop comments and string bodies, keep numbering."""
    raw, joined, buf, start = text.split("\n"), [], "", 1
    for i, line in enumerate(raw, 1):
        if not buf:
            start = i
        stripped = line.rstrip()
        if stripped.endswith(" _"):
            buf += stripped[:-1]
            continue
        buf += stripped
        joined.append((start, buf))
        buf = ""
    if buf:
        joined.append((start, buf))
    clean = []
    for n, line in joined:
        out, in_str = [], False
        k = 0
        while k < len(line):
            ch = line[k]
            if ch == '"':
                if in_str and k + 1 < len(line) and line[k + 1] == '"':
                    k += 2
                    continue
                in_str = not in_str
                out.append('"')
            elif not in_str and ch == "'":
                break
            else:
                out.append(" " if in_str else ch)
            k += 1
        s = "".join(out).strip()
        if re.match(r"(?i)^rem\b", s):
            s = ""
        clean.append((n, s))
    return clean

OPEN = [
    (re.compile(r"(?i)^(?:(?:public|private|friend|static)\s+)*(?:sub|function|property\s+(?:get|let|set))\s+\w+"), "proc"),
    (re.compile(r"(?i)^with\b"), "with"),
    (re.compile(r"(?i)^select\s+case\b"), "select"),
    (re.compile(r"(?i)^do\b"), "do"),
    (re.compile(r"(?i)^for\b"), "for"),
]
CLOSE = [
    ("proc",   re.compile(r"(?i)^end\s+(?:sub|function|property)\b")),
    ("with",   re.compile(r"(?i)^end\s+with\b")),
    ("select", re.compile(r"(?i)^end\s+select\b")),
    ("do",     re.compile(r"(?i)^loop\b")),
    ("for",    re.compile(r"(?i)^next\b")),
]

def statements(line):
    """VBA lets several statements share a line. Split on ':', but a bare
    'Label:' is not a separator."""
    if re.match(r"^[A-Za-z]\w*:\s*$", line):
        return [line]
    return [part.strip() for part in line.split(":") if part.strip()]

def check_blocks(name, lines, problems):
    stack = []
    for n, raw in lines:
        if not raw:
            continue
        for s in statements(raw):
            if re.match(r"(?i)^(declare|type|enum|attribute|option|const|dim|redim|exit)\b", s):
                continue
            if re.match(r"(?i)^end\s+(type|enum)\b", s):
                continue
            if re.match(r"^[A-Za-z]\w*:\s*$", s):
                continue
            m = re.match(r"(?i)^(if|elseif)\b(.*)$", s)
            if m:
                rest = m.group(2)
                tail = re.split(r"(?i)\bthen\b", rest, maxsplit=1)
                if m.group(1).lower() == "if" and len(tail) > 1 and not tail[1].strip():
                    stack.append(("if", n))
                continue
            if re.match(r"(?i)^else\b", s):
                continue
            if re.match(r"(?i)^end\s+if\b", s):
                if not stack or stack[-1][0] != "if":
                    problems.append(f"{name}:{n}  End If without a matching multi-line If")
                else:
                    stack.pop()
                continue
            opened = False
            for rx, kind in OPEN:
                if rx.match(s):
                    stack.append((kind, n)); opened = True; break
            if opened:
                continue
            for kind, rx in CLOSE:
                if rx.match(s):
                    if not stack:
                        problems.append(f"{name}:{n}  '{s[:24]}' closes nothing")
                    elif stack[-1][0] != kind:
                        problems.append(f"{name}:{n}  '{s[:24]}' closes a {kind}, but the open block is {stack[-1][0]} from line {stack[-1][1]}")
                        stack.pop()
                    else:
                        stack.pop()
                    break
    for kind, n in stack:
        problems.append(f"{name}:{n}  {kind} block is never closed")

# Anything here can move data off the machine. The product must contain none of
# it. This is the one property a corporate data policy actually cares about, so
# it is enforced at build time rather than asserted in a document: adding any of
# these to word/ fails the build.
FORBIDDEN = [
    (r"(?i)\bMSXML2\.(?:Server)?XMLHTTP",     "HTTP client"),
    (r"(?i)\bWinHttp\.",                      "HTTP client"),
    (r"(?i)\bInternetExplorer\.Application", "browser automation"),
    (r"(?i)\bMSXML2\.XMLHTTP",                "HTTP client"),
    (r"(?i)\bADODB\.Connection\b",           "database or remote connection"),
    (r"(?i)\bCDO\.(?:Message|Configuration)", "email"),
    (r"(?i)\bOutlook\.Application\b",        "email"),
    (r"(?i)\.SendMail\b",                    "email"),
    (r"(?i)\bMailEnvelope\b",                "email"),
    (r"(?i)\bSendFax\b",                     "fax"),
    (r"(?i)\bFollowHyperlink\b",             "opens a URL"),
    (r"(?i)\bPublishObjects\b",              "publishes to a location"),
    (r"(?i)\bWebService\s*\(",               "web request"),
    (r"(?i)\bURLDownloadToFile\b",           "download"),
    (r"(?i)\bwininet|urlmon\b",              "network library"),
    (r"(?i)\bWinsock\b",                     "sockets"),
    (r"(?i)\bcertutil\b.*-urlcache",         "download via certutil"),
    (r"(?i)\b(?:bitsadmin|curl|wget|Invoke-WebRequest|Invoke-RestMethod)\b", "download command"),
    (r"(?i)resolveExternals\s*=\s*True",      "would let a crafted file make the XML parser fetch"),
    (r"(?i)ProhibitDTD\"?\s*,\s*False",       "would let a crafted file make the XML parser fetch"),
    (r"(?i)Target(?:Mode)?\s*=\s*\"?External", "external relationship"),
]
# http(s) in the source is only ever an XML namespace, which is an identifier
# and is never fetched. Anything else is a finding.
URL_OK = re.compile(r"(?i)https?://schemas\.(?:openxmlformats\.org|microsoft\.com)/")

# RFC 2606 reserves example.com/net/org so they can never resolve to anything
# real. The self-check needs them as the addresses it proves are REFUSED, and
# nowhere else in the product may mention an address at all.
EXAMPLE_ONLY_IN = "CLSelfCheck.bas"
EXAMPLE = re.compile(r"(?i)^[a-z][a-z0-9+.-]*://(?:[a-z0-9-]+\.)*example\.(?:com|net|org)(?:[/:?#]|$)")


def check_no_egress(mods, builder, problems):
    # Scanned RAW, not through logical_lines: an address or a COM class name
    # lives inside a string literal, and comments can hide a command too.
    for name, text in list(mods.items()) + [("build/CLBuild.bas", builder)]:
        for n, line in enumerate(text.split("\n"), 1):
            if not line.strip():
                continue
            for pattern, why in FORBIDDEN:
                if re.search(pattern, line):
                    problems.append(f"{name}:{n}  contains {why} ({pattern}) - the product must not be able to send data anywhere")
            for url in re.findall(r"(?i)\b[a-z][a-z0-9+.-]*://[^\s\"'<>)]+", line):
                if URL_OK.match(url):
                    continue
                if EXAMPLE.match(url) and name == EXAMPLE_ONLY_IN:
                    continue
                problems.append(f"{name}:{n}  refers to {url} - the product must not reach any address")
    # certutil is used by the builder only, and only to hash a local file with
    # -hashfile. It is allowed there, named explicitly, and never ships.
    for name, text in mods.items():
        if re.search(r"(?i)\bcertutil\b", text):
            problems.append(f"{name}  mentions certutil; it belongs in the builder only, never in shipped code")
    for n, line in enumerate(builder.split("\n"), 1):
        if re.search(r"(?i)\bcertutil\b", line) and "-hashfile" not in line and not line.lstrip().startswith("'"):
            problems.append(f"build/CLBuild.bas:{n}  uses certutil for something other than -hashfile")


SIG = re.compile(r"(?i)^\s*(?:(?:public|private|friend)\s+)?(?:static\s+)?(sub|function)\s+(\w+)\s*\((.*)\)\s*(?:as\s+\w+)?\s*$")

def split_top(text):
    """Split on commas that are not inside brackets or quotes."""
    out, depth, buf = [], 0, ""
    for ch in text:
        if ch in "([": depth += 1
        elif ch in ")]": depth -= 1
        if ch == "," and depth == 0:
            out.append(buf); buf = ""
        else:
            buf += ch
    if buf.strip(): out.append(buf)
    return [p.strip() for p in out if p.strip()]

def param_bounds(params):
    parts = split_top(params)
    required = sum(1 for p in parts if not re.match(r"(?i)^\s*optional\b", p)
                   and not re.match(r"(?i)^\s*paramarray\b", p))
    if any(re.match(r"(?i)^\s*paramarray\b", p) for p in parts):
        return required, 99
    return required, len(parts)

def check_arity(mods, problems):
    """A compiler would reject a call with the wrong number of arguments. This
    only inspects unambiguous statement-position calls, so it never guesses."""
    sigs = {}
    for name, text in mods.items():
        for n, line in logical_lines(text):
            m = SIG.match(line)
            if m:
                sigs[m.group(2).lower()] = (m.group(2), param_bounds(m.group(3)), name)
    for name, text in mods.items():
        for n, line in logical_lines(text):
            if not line or SIG.match(line):
                continue
            pieces = [s.strip() for s in line.split(":") if s.strip()]
            # a one-line "If x Then DoThing a, b" hides a call after Then
            for piece in list(pieces):
                tail = re.split(r"(?i)\bthen\b", piece, maxsplit=1)
                if len(tail) > 1 and tail[1].strip():
                    pieces.append(tail[1].strip())
                tail = re.split(r"(?i)\belse\b", piece, maxsplit=1)
                if len(tail) > 1 and tail[1].strip():
                    pieces.append(tail[1].strip())
            for stmt in pieces:
                stmt = re.sub(r"(?i)^call\s+", "", stmt)
                m = re.match(r"^(CL\w+)\s*$", stmt)
                if m and m.group(1).lower() in sigs:
                    proc, (lo, hi), where = sigs[m.group(1).lower()]
                    if lo > 0:
                        problems.append(f"{name}:{n}  {proc} is called with no arguments but needs {lo} ({where})")
                    continue
                m = re.match(r"^(CL\w+)\s+(?!=)(.+)$", stmt)
                if not m or m.group(1).lower() not in sigs:
                    continue
                if re.match(r"(?i)^(as|is|then|to|=)\b", m.group(2)):
                    continue
                proc, (lo, hi), where = sigs[m.group(1).lower()]
                count = len(split_top(m.group(2)))
                if count < lo or count > hi:
                    problems.append(f"{name}:{n}  {proc} is called with {count} argument(s); it takes "
                                    f"{lo if lo == hi else str(lo) + ' to ' + str(hi)} ({where})")

def main():
    mods, problems, notes = modules(), [], []
    builder = (BUILD / "CLBuild.bas").read_text(encoding="utf-8")
    defined, public, calls, labels = {}, {}, [], []

    # A .cls file defines a type named after itself, and opens with a header
    # block that is not VBA statements. Strip exactly that block - "END" alone,
    # never "End Sub".
    for name in list(mods):
        if not name.endswith(".cls"):
            continue
        defined.setdefault(name[:-4].lower(), []).append(name)
        kept, in_header = [], False
        for line in mods[name].split("\n"):
            bare = line.strip()
            if re.match(r"(?i)^VERSION\s", bare):
                continue
            if bare.upper() == "BEGIN":
                in_header = True
                continue
            if in_header:
                if bare.upper() == "END":
                    in_header = False
                continue
            if re.match(r"(?i)^Attribute\s", bare):
                continue
            kept.append(line)
        mods[name] = "\n".join(kept)

    for name, text in mods.items():
        lines = logical_lines(text)
        if "Option Explicit" not in text:
            problems.append(f"{name}  is missing Option Explicit")
        for n, line in enumerate(text.split("\n"), 1):
            if len(line) > 1023:
                problems.append(f"{name}:{n}  line is {len(line)} characters; VBA allows 1023")
        check_blocks(name, lines, problems)
        for n, s in lines:
            m = re.match(r"(?i)^(?:(public|private|friend)\s+)?(?:static\s+)?(sub|function|property\s+\w+)\s+(\w+)", s)
            if m:
                proc = m.group(3)
                defined.setdefault(proc.lower(), []).append(name)
                if (m.group(1) or "public").lower() != "private" and name.endswith(".bas"):
                    public.setdefault(proc.lower(), []).append((proc, name))
            m = re.match(r"^([A-Za-z]\w*):\s*$", s)
            if m:
                labels.append((name, n, m.group(1)))
            for tok in re.findall(r"\b(CL[A-Za-z0-9_]*)\b", s):
                calls.append((name, n, tok))

    # constants, module variables and enum-ish public data also count as defined names
    for name, text in mods.items():
        for m in re.finditer(r"(?im)^\s*(?:public|private|global)\s+(?:const\s+)?(\w+)", text):
            defined.setdefault(m.group(1).lower(), []).append(name)
        for m in re.finditer(r"(?im)^\s*public\s+const\s+(\w+)", text):
            defined.setdefault(m.group(1).lower(), []).append(name)

    for proc, where in public.items():
        if len({w[1] for w in where}) > 1:
            problems.append(f"{where[0][0]}  is Public in more than one module: {', '.join(sorted({w[1] for w in where}))}")

    RESERVED = {"empty", "error", "next", "stop", "end", "name", "close", "open", "print",
                "input", "line", "get", "put", "loop", "case", "is", "like", "mod", "not",
                "and", "or", "new", "nothing", "null", "true", "false", "me", "option",
                "resume", "return", "select", "static", "sub", "function", "set", "let", "to", "then"}
    for name, n, lab in labels:
        if lab.lower() in RESERVED:
            problems.append(f"{name}:{n}  '{lab}' is a VBA keyword and cannot be a line label")

    unknown = {}
    for name, n, tok in calls:
        if tok.lower() in defined or tok.lower() in BUILTIN:
            continue
        if re.match(r"(?i)^cl(ng|ose|ear|ick|s)$", tok):
            continue
        unknown.setdefault(tok, []).append(f"{name}:{n}")
    for tok, where in sorted(unknown.items()):
        problems.append(f"{tok}  is used but never defined ({where[0]}{', +%d more' % (len(where)-1) if len(where) > 1 else ''})")

    check_arity(mods, problems)

    # ribbon callbacks must exist, and take exactly one argument
    for seed, label in (("seed/ClauseLibraryRuntime.dotm", "runtime"), ("seed/Start Here.docm", "setup")):
        import zipfile
        from xml.etree import ElementTree as ET
        path = BUILD / seed
        if not path.exists():
            notes.append(f"seed {label} not built yet - skipping its ribbon check")
            continue
        xml = zipfile.ZipFile(path).read("customUI/customUI14.xml").decode()
        ET.fromstring(xml)
        for cb in sorted(set(re.findall(r'(?:onAction|getContent|onLoad)="([^"]+)"', xml))):
            hit = None
            for name, text in mods.items():
                m = re.search(r"(?im)^\s*public\s+(sub|function)\s+" + re.escape(cb) + r"\s*\(([^)]*)\)", text)
                if m:
                    hit = (name, m.group(2))
                    break
            if not hit:
                problems.append(f"ribbon ({label}) calls {cb}, which no module defines")
            elif hit[1].count(",") != 0 or not hit[1].strip():
                problems.append(f"ribbon ({label}) callback {cb} must take exactly one argument; it takes '{hit[1].strip()}'")

    # every control the form code touches must be created by the builder
    per_form = {}
    for m in re.finditer(r"Build(\w+)Form\b.*?(?=Private Sub Build|\Z)", builder, re.S):
        block = m.group(0)
        key = "frm" + m.group(1)
        per_form[key] = set(re.findall(r'Add designer, "\w+", "(\w+)"', block))
    CONTROL = r"\b((?:lbl|txt|cbo|cmd|lst|chk)[A-Za-z0-9]+)\b"
    for form, src in (("frmLibrary", "LibraryForm.txt"), ("frmHistory", "HistoryForm.txt")):
        if src not in mods:
            continue
        made = per_form.get(form, set())
        used = set()
        for n, line in logical_lines(mods[src]):
            used |= set(re.findall(CONTROL, line))
            used |= set(re.findall(r'Me\.Controls\("(\w+)"\)', line))
        for c in sorted(used - made):
            problems.append(f"{src}  uses control '{c}', which the builder never creates for {form}")
        unused = made - used
        if unused:
            notes.append(f"{form}: created but never referenced in code: {', '.join(sorted(unused))}")

    # handlers must match a control that exists
    for src in ("LibraryForm.txt", "HistoryForm.txt"):
        if src not in mods:
            continue
        for m in re.finditer(r"(?im)^\s*private\s+sub\s+(\w+)_(Click|Change|DblClick|KeyDown|Enter|Exit)\b", mods[src]):
            form = "frmLibrary" if src.startswith("Library") else "frmHistory"
            if m.group(1) not in per_form.get(form, set()) and m.group(1) != "UserForm":
                problems.append(f"{src}  has a handler for '{m.group(1)}', which is not a control the builder creates")

    check_no_egress(mods, builder, problems)

    for cls in sorted(WORD.glob("*.cls")):
        if f'ImportClass fso, d, "{cls.stem}"' not in builder:
            problems.append(f"{cls.name}  exists but the builder never imports it into the template")

    # Each shipped file gets its own VBA project with its own module set. A
    # call that resolves across the whole source tree can still fail to compile
    # in the smaller setup project, which carries only three modules and a shim.
    PROJECTS = {
        "ClauseLibraryPersonal.dotm": (
            ["CLPlatform", "CLStore", "CLRich", "CLActions", "CLExport",
             "CLRecovery", "CLSetup", "CLSelfCheck"],
            ["LibraryForm.txt", "HistoryForm.txt", "CLDraftWatcher.cls"], set()),
        "Start Here.docm": (
            ["CLPlatform", "CLStore", "CLSetup"], [],
            {"clclosemanager", "clreleasemanager", "clensureready",
             "clnotifymanager", "clrefreshribbon"}),
    }
    shim_source = re.search(r'shim\.CodeModule\.AddFromString(.*?)\n\s*d\.Save', builder, re.S)
    shim_names = set(re.findall(r'(?i)Public (?:Sub|Function) (\w+)', shim_source.group(1))) if shim_source else set()
    for target, (basNames, forms, declared) in PROJECTS.items():
        if declared and {n.lower() for n in shim_names} != declared:
            problems.append(f"{target}: the builder's shim defines {sorted(shim_names)}, "
                            f"but this check expects {sorted(declared)} - keep them in step")
        files = [n + ".bas" for n in basNames] + forms
        available = set(declared)
        for f in files:
            if f not in mods:
                problems.append(f"{target}: source file {f} is missing")
                continue
            if f.endswith(".cls"):
                available.add(f[:-4].lower())
            for m in re.finditer(r"(?im)^\s*(?:public|private)\s+(?:const\s+)?(\w+)", mods[f]):
                available.add(m.group(1).lower())
            for m in re.finditer(r"(?im)^\s*(?:public|private|friend)?\s*(?:static\s+)?"
                                 r"(?:sub|function)\s+(\w+)", mods[f]):
                available.add(m.group(1).lower())
        missing = {}
        for f in files:
            if f not in mods:
                continue
            for n, line in logical_lines(mods[f]):
                for tok in re.findall(r"\b(CL[A-Za-z0-9_]*)\b", line):
                    if tok.lower() not in available and tok.lower() not in BUILTIN \
                       and not re.match(r"(?i)^cl(ng|ose|ear|ick|s)$", tok):
                        missing.setdefault(tok, f"{f}:{n}")
        for tok, where in sorted(missing.items()):
            problems.append(f"{target}: {tok} is used at {where} but no module in that project defines it")

    print(f"Checked {len(mods)} source files, {sum(len(v) for v in defined.values())} definitions, "
          f"{len(PROJECTS)} projects.\n")
    for note in notes:
        print(f"  note: {note}")
    if problems:
        print(f"\n{len(problems)} problem(s):\n")
        for p in problems:
            print(f"  {p}")
        return 1
    print("\nAll static checks passed.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
