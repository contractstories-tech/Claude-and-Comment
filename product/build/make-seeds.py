#!/usr/bin/env python3
"""Build the two seed packages the Word builder fills in.

A seed carries everything Word will NOT write for us: the ribbon definition,
the content types and relationships, and (for Start Here) the page the user
reads. It also carries a known-good macro project, purely so the file opens;
the builder replaces every module in it.

Run:  python3 build/make-seeds.py <folder containing the 1.0 package>
"""
import re, shutil, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

HERE = Path(__file__).resolve().parent
SEED = HERE / "seed"
UI   = "http://schemas.microsoft.com/office/2009/07/customui"
REL  = "http://schemas.openxmlformats.org/package/2006/relationships"
W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

RUNTIME_RIBBON = f"""<customUI xmlns="{UI}" onLoad="CLRibbonLoad">
<ribbon><tabs><tab id="CLTab" label="Clause Library">
 <group id="CLDraft" label="Drafting">
  <dynamicMenu id="CLInsertMenu" label="Insert" size="large" imageMso="PasteTextOnly"
     getContent="CLQuickMenu"
     screentip="Insert wording you have used recently"
     supertip="Your most recently used clauses and comments, ready to drop in where the cursor is."/>
  <button id="CLCapture" label="Capture selection" size="large" imageMso="Copy" onAction="CLRibbonCapture"
     screentip="Keep the selected wording"
     supertip="Saves a reusable copy. This document is never changed."/>
  <button id="CLOpen" label="Library" size="large" imageMso="FileOpen" onAction="CLRibbonOpen"
     screentip="Search everything you have kept"
     supertip="Find wording, read it, insert it, or organise it. New captures are usable straight away."/>
 </group>
 <group id="CLKeep" label="Your library">
  <dynamicMenu id="CLFavouritesMenu" label="Favourites" imageMso="ReviewHighlightChanges" getContent="CLQuickMenu"/>
  <button id="CLCaptureComment" label="Capture comment" imageMso="ReviewNewComment" onAction="CLRibbonCaptureComment"
     screentip="Save a comment you want to reuse"/>
  <button id="CLSaveWording" label="Save wording" imageMso="FileSave" onAction="CLRibbonSave"
     screentip="Put an edited draft back into the library"
     supertip="Use this in a draft opened with Edit wording. The version you replace is always kept."/>
 </group>
 <group id="CLCare" label="Looking after it">
  <button id="CLBackup" label="Back up" imageMso="SaveAll" onAction="CLRibbonBackup"/>
  <button id="CLExport" label="Export" imageMso="ExportSavedExports" onAction="CLRibbonExport"
     screentip="Write your clauses out as ordinary Word files"/>
  <button id="CLGuide" label="Quick guide" imageMso="Help" onAction="CLRibbonHelp"/>
  <menu id="CLOptions" label="Options" imageMso="ControlProperties">
   <button id="CLRestoreBackup" label="Restore a backup..." onAction="CLRibbonRestoreBackup"/>
   <button id="CLLocate" label="Locate a library..." onAction="CLRibbonLocate"/>
   <menuSeparator id="CLOptSep1"/>
   <button id="CLSelfCheck" label="Run self-check" imageMso="RecordsCheckIn" onAction="CLRibbonSelfCheck"
      screentip="Check this installation is working"
      supertip="Runs about eighty checks against a throwaway library. Your own library is not touched."/>
   <button id="CLAbout" label="About Clause Library" onAction="CLRibbonAbout"/>
   <menuSeparator id="CLOptSep2"/>
   <button id="CLRemove" label="Remove from Word..." imageMso="Delete" onAction="CLRibbonRemove"/>
  </menu>
 </group>
</tab></tabs></ribbon></customUI>"""

SETUP_RIBBON = f"""<customUI xmlns="{UI}">
<ribbon><tabs><tab id="CLSetupTab" label="Clause Library Setup">
 <group id="CLSetupGroup" label="One-time setup">
  <button id="CLInstall" label="Set up Clause Library" size="large" imageMso="FileOpen" onAction="CLRibbonSetup"
     screentip="Add Clause Library to Word"
     supertip="Puts the tools on a Clause Library tab in every document, and creates your library folder. No administrator rights needed."/>
  <button id="CLSetupLocate" label="Use an existing library..." imageMso="FolderOpen" onAction="CLRibbonSetupLocate"
     screentip="Point at a library folder you already have"/>
 </group>
 <group id="CLRemoveGroup" label="Removing it">
  <button id="CLUninstall" label="Remove from Word" size="large" imageMso="Delete" onAction="CLRibbonRemoveHere"
     screentip="Take the tools out and keep every clause"
     supertip="Your library, previous versions and backups all stay exactly where they are."/>
  <button id="CLRemoveAll" label="Remove everything, including my clauses..." imageMso="Cancel" onAction="CLRibbonRemoveEverything"
     screentip="Delete Clause Library and all its data from this computer"/>
 </group>
</tab></tabs></ribbon></customUI>"""

PAGE = [
 ("Title",    "Clause Library"),
 ("Normal",   "Keep the wording you trust where you actually draft. This is a personal library "
              "that works offline, in desktop Word for Windows. Nothing is sent anywhere."),
 ("Heading2", "If Word says this file came from the internet"),
 ("Normal",   "Close Word. In File Explorer, right-click this document, choose Properties, tick "
              "Unblock at the bottom, and click OK. Do the same for ClauseLibraryPersonal.dotm "
              "next to it. Then open this document again. Windows blocks macros in files that "
              "arrived by email or download until you do this."),
 ("Heading2", "Set it up once"),
 ("Normal",   "Keep this document and ClauseLibraryPersonal.dotm together in the folder you "
              "extracted. Open the Clause Library Setup tab above and choose Set up Clause "
              "Library. Word needs to allow macros; nothing else is installed and no "
              "administrator rights are needed."),
 ("Heading2", "Check it works"),
 ("Normal",   "On the Clause Library tab in any document, choose Options, then Run self-check. "
              "It runs about eighty checks against a throwaway library and tells you plainly "
              "whether everything passed. Your own library is never touched."),
 ("Heading2", "Then, in any document"),
 ("Normal",   "Select wording worth keeping and choose Capture selection. Give it a name when "
              "asked - that is what you will search for later. Your document is never changed."),
 ("Normal",   "To use something again, put the cursor where it should go and choose Insert for "
              "what you used recently, or Library to search everything. Insertion follows "
              "whatever Track Changes setting that document is using."),
 ("Heading2", "Keeping it safe"),
 ("Normal",   "Back up writes a checked copy wherever you choose - use a different drive, so "
              "that losing this computer does not lose your library. Export writes every clause "
              "out as ordinary Word files that need no macros and no Clause Library to read. "
              "Every version of every item is kept, and nothing is ever deleted outright."),
 ("Heading2", "Removing it"),
 ("Normal",   "Close the library window, then choose Remove from Word on this tab. Your clauses, "
              "their previous versions and your backups all stay where they are. There is also a "
              "complete removal, which deletes your clauses too - it asks you to type the word "
              "DELETE first."),
]

def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def body_xml(original: str) -> str:
    head = original[:original.index("<w:body>") + len("<w:body>")]
    sect = original[original.index("<w:sectPr"):]
    paras = []
    for style, text in PAGE:
        pr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style != "Normal" else \
             '<w:pPr><w:spacing w:after="160"/></w:pPr>'
        paras.append(f'<w:p>{pr}<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')
    return head + "".join(paras) + sect

def patch(src: Path, dst: Path, ribbon: str, new_body: bool):
    with zipfile.ZipFile(src) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    ET.fromstring(ribbon)                                   # refuse to ship malformed ribbon XML
    ids = re.findall(r'\sid="([^"]+)"', ribbon)
    assert len(ids) == len(set(ids)), f"duplicate ribbon ids: {[i for i in ids if ids.count(i) > 1]}"
    parts["customUI/customUI14.xml"] = ribbon.encode("utf-8")

    rels = ET.fromstring(parts["_rels/.rels"])
    for old in list(rels):
        if old.get("Type", "").endswith("/ui/extensibility"):
            rels.remove(old)
    r = ET.SubElement(rels, f"{{{REL}}}Relationship")
    r.set("Id", "rIdClauseLibraryRibbon")
    r.set("Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility")
    r.set("Target", "customUI/customUI14.xml")
    parts["_rels/.rels"] = ET.tostring(rels, xml_declaration=True, encoding="UTF-8")

    if new_body:
        parts["word/document.xml"] = body_xml(parts["word/document.xml"].decode("utf-8")).encode("utf-8")

    for name in ("docProps/core.xml", "docProps/app.xml"):
        if name in parts:
            props = ET.fromstring(parts[name])
            for item in list(props):
                if item.tag.rsplit("}", 1)[-1] in ("creator", "lastModifiedBy", "Company", "Manager", "HyperlinkBase"):
                    props.remove(item)
            parts[name] = ET.tostring(props, xml_declaration=True, encoding="UTF-8")

    for name in [n for n in parts if n.endswith(".rels")]:
        rs = ET.fromstring(parts[name])
        for rr in list(rs):
            if rr.get("Type", "").endswith("/attachedTemplate"):
                rs.remove(rr)
        parts[name] = ET.tostring(rs, xml_declaration=True, encoding="UTF-8")

    assert "word/vbaProject.bin" in parts, "seed must carry a macro project so Word will open it"
    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in parts.items():
            z.writestr(n, b)
    with zipfile.ZipFile(dst) as z:                          # prove the result is a sane package
        assert z.testzip() is None
        for n in z.namelist():
            if n.endswith((".xml", ".rels")):
                ET.fromstring(z.read(n))
    return dst

def main():
    if len(sys.argv) < 2:
        sys.exit("usage: make-seeds.py <folder holding ClauseLibrary 1.0>")
    src = Path(sys.argv[1])
    a = patch(src / "ClauseLibraryPersonal.dotm", SEED / "ClauseLibraryRuntime.dotm", RUNTIME_RIBBON, False)
    b = patch(src / "Start Here.docm",            SEED / "Start Here.docm",            SETUP_RIBBON,   True)
    for p in (a, b):
        print(f"  wrote {p.name:32} {p.stat().st_size:,} bytes")

if __name__ == "__main__":
    main()
