Attribute VB_Name = "CLSelfCheck"
' Clause Library - the self-check.
'
' Everything here runs against a throwaway library in your temporary folder and
' throwaway documents. Your own library is never opened, read or written.
' Run it after installing, after a Word or Windows update, or any time
' something behaves oddly - and send the log if you need help.
Option Explicit

Private passed As Long
Private failed As Long
Private log As String
Private savedRoot As String

Private Sub Check(ByVal condition As Boolean, ByVal name As String)
    If condition Then
        passed = passed + 1: log = log & "PASS  " & name & vbCrLf
    Else
        failed = failed + 1: log = log & "FAIL  " & name & vbCrLf
    End If
End Sub

Private Sub Note(ByVal text As String)
    log = log & vbCrLf & "-- " & text & " " & String$(CLMax(2, 62 - Len(text)), "-") & vbCrLf
End Sub

Private Function CLMax(ByVal a As Long, ByVal b As Long) As Long
    CLMax = a: If b > a Then CLMax = b
End Function

Public Sub CLRibbonSelfCheck(ByVal control As Object): CLRunSelfCheck: End Sub

Public Sub CLRunSelfCheck()
    Dim root As String, logPath As String, started As Date
    started = Now
    passed = 0: failed = 0: log = ""
    savedRoot = CLTestRoot
    root = Environ$("TEMP") & "\ClauseLibrarySelfCheck-" & Left$(CLId(), 8)
    logPath = Environ$("TEMP") & "\Clause Library self-check.txt"
    log = "CLAUSE LIBRARY SELF-CHECK" & vbCrLf & _
          "Version " & CL_VERSION & " | library format " & CL_SCHEMA & vbCrLf & _
          "Word " & Application.Version & " " & Application.Build & vbCrLf & _
          "Run " & CLStamp(started) & vbCrLf & _
          "Throwaway library: " & root & vbCrLf & String$(68, "=") & vbCrLf

    Application.ScreenUpdating = False
    On Error GoTo Fatal
    CLCreateLibrary root
    CLTestRoot = root
    CheckPlatform
    CheckStorage
    CheckSearch
    CheckRecovery
    CheckWord
    GoTo Finish
Fatal:
    failed = failed + 1
    log = log & vbCrLf & "STOPPED  The check could not finish: " & CLExplain(Err.number, Err.Description) & vbCrLf
Finish:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.StatusBar = False
    CLTestRoot = savedRoot
    CLForgetCache
    log = log & String$(68, "=") & vbCrLf & passed & " passed, " & failed & " failed" & vbCrLf & _
          "Finished in " & DateDiff("s", started, Now) & " seconds." & vbCrLf
    If CLFolderExists(root) Then CLFso().DeleteFolder root, True
    CLRequireLocalPath logPath, "The log"
    CLWrite logPath, log
    On Error GoTo 0
    Dim headline As String
    If failed = 0 Then
        headline = "All " & passed & " checks passed." & vbCrLf & vbCrLf & _
                   "Clause Library is working correctly on this computer, with this version of Word."
    Else
        headline = passed & " checks passed, " & failed & " FAILED." & vbCrLf & vbCrLf & _
                   "Do not rely on this installation until the failures are understood. The log lists exactly what failed."
    End If
    MsgBox headline & vbCrLf & vbCrLf & "A full log was saved to:" & vbCrLf & logPath & vbCrLf & vbCrLf & _
           "Your own library was not touched.", IIf(failed = 0, vbInformation, vbExclamation), "Self-check"
    On Error Resume Next
    If CLIsLocalPath(logPath) Then Documents.Open logPath
End Sub

' ---------- platform ----------

Private Sub CheckPlatform()
    Note "Checksums and text safety"
    Check CLSha("") = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855", "SHA-256 of nothing matches the published value"
    Check CLSha("abc") = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "SHA-256 of 'abc' matches the published value"
    Check CLSha("clause") <> CLSha("clauses"), "Different wording gives a different checksum"

    ' The characters Word puts in Range.Text that XML cannot hold. Before this
    ' was handled, capturing a clause with a footnote or a non-breaking hyphen
    ' failed with a message about library files.
    Dim awkward As String
    awkward = "Clause" & Chr$(2) & " with a footnote, a" & Chr$(30) & "non-breaking hyphen," & _
              Chr$(31) & " a soft hyphen, an image" & Chr$(1) & ", a page" & Chr$(12) & "break and a cell" & Chr$(7) & "end."
    Dim safe As String: safe = CLSafeText(awkward)
    Check InStr(safe, Chr$(1)) = 0 And InStr(safe, Chr$(2)) = 0 And InStr(safe, Chr$(31)) = 0, "Illegal control characters are removed"
    Check InStr(safe, "non-breaking") > 0 And InStr(safe, "footnote") > 0, "The wording itself survives that cleaning"
    Dim probe As Object
    Set probe = CLNew()
    CLSet probe, "plain", awkward
    Check CLDom(probe.XML).documentElement.nodeName = "entry", "A record holding that wording is still valid XML"

    Check CLStamp(DateSerial(2026, 3, 7) + TimeSerial(9, 4, 5)) = "2026-03-07T09:04:05", "Timestamps ignore regional settings"
    Check CLNiceDate("2026-09-19T14:02:33") = "19 Sep 2026", "Dates are shown in a readable form"
    Check CLIsId(CLId()) And Not CLIsId("not-an-id"), "Identifiers are generated and validated"
    Check CLId() <> CLId(), "Every identifier is different"
    Check CLSafeFileName("Limitation of liability: cap / carve-outs?", 70) = "Limitation of liability cap carve-outs", "File names are made safe for Windows"

    Note "Keeping data on this machine"
    ' Word will open a web address as readily as a file, and explorer.exe will
    ' launch a browser. Every path this product hands to either of them goes
    ' through this first.
    Check CLIsLocalPath("C:\\Users\\me\\Clauses"), "An ordinary folder is accepted"
    Check CLIsLocalPath("\\\\fileserver\\legal\\Clauses"), "A network share on your own network is accepted"
    Check Not CLIsLocalPath("https://example.com/clauses"), "A web address is refused"
    Check Not CLIsLocalPath("http://example.com/clauses"), "An insecure web address is refused"
    Check Not CLIsLocalPath("ftp://example.com/clauses"), "Any other internet address is refused"
    Check Not CLIsLocalPath("file://example.com/share"), "Even a file scheme is refused - only real paths are used"
    Dim refused As Boolean
    On Error Resume Next
    CLRequireLocalPath "https://example.com/x", "Test"
    refused = (Err.number <> 0): Err.Clear
    On Error GoTo 0
    Check refused, "Asking for a web address stops the operation"

    ' The product sends nothing anywhere. A folder inside a sync service does,
    ' which is the one route out that it can warn about but not prevent.
    Check CLSyncService("C:\\Users\\me\\AppData\\Local\\ClauseLibraryPersonal") = "", "An ordinary local folder is not flagged"
    Check CLSyncService("C:\\Users\\me\\OneDrive - Acme LLP\\Clauses") = "OneDrive", "A OneDrive folder is recognised"
    Check CLSyncService("C:\\Users\\me\\Dropbox\\Clauses") = "Dropbox", "A Dropbox folder is recognised"
    Check CLSyncService("D:\\Google Drive\\Precedents") = "Google Drive", "A Google Drive folder is recognised"
    Check CLSyncService("\\\\fileserver\\legal\\Clauses") = "", "A folder on your own file server is not a sync service"
End Sub

' ---------- storage ----------

Private Sub CheckStorage()
    Note "Saving, reloading and protecting records"
    Dim e As Object, loaded As Object, status As String, id As String
    Set e = CLNew()
    CLSet e, "title", "Mutual indemnity"
    CLSet e, "plain", "Each party shall indemnify the other against all Losses."
    CLSet e, "notes", "Only I see this."
    CLSave e, 0, "<payload/>", True
    id = CLEntryId(e)
    Set loaded = CLLoad(id, status)
    Check CLRevision(loaded) = 1 And status = "ok", "A saved item reloads unchanged"
    Check CLGet(loaded, "title") = "Mutual indemnity", "Details survive the round trip"
    Check CLExists(CLPayloadPath(id)), "The Word payload is stored in its own file"
    Check CLReadPayload(loaded) = "<payload/>", "The payload reads back and its checksum verifies"
    Check CLFso().GetFile(CLEntryPath(id)).Size < 4000, "The record itself stays small, so listing stays fast"

    CLSet loaded, "topic", "Liability": CLSave loaded, 1
    Check CLRevision(CLLoad(id)) = 2, "Changing details makes a new version"
    Check CLExists(CLRoot() & "\history\" & id & "-r0000000001.xml"), "The version it replaced is kept"

    Dim blocked As Boolean
    On Error Resume Next
    CLSave e, 1: blocked = (Err.number <> 0): Err.Clear
    On Error GoTo 0
    Check blocked, "A window holding an out-of-date copy cannot overwrite a newer one"
    Check CLRevision(CLLoad(id)) = 2, "That refusal left the saved item exactly as it was"

    Dim gate As Integer, locked As Boolean
    gate = FreeFile
    Open CLRoot() & "\writer.lock" For Binary Access Read Write Lock Read Write As #gate
    On Error Resume Next
    Dim other As Object: Set other = CLLoad(id)
    CLSet other, "tags", "blocked"
    CLSave other, CLRevision(other): locked = (Err.number <> 0): Err.Clear
    On Error GoTo 0
    Close #gate
    Check locked, "A second writer is held off while another is saving"
    Check CLExplain(70, "Permission denied") <> "Permission denied", "Being held off is explained in plain words"
    Check CLExplain(5941, "x") <> "x", "Using Save wording on the wrong document is explained in plain words"

    ' Tampering is noticed, but the record stays visible and repairable rather
    ' than disappearing from the library.
    Dim tampered As Object: Set tampered = CLLoad(id)
    CLWrite CLEntryPath(id), Replace$(tampered.XML, "Mutual indemnity", "Edited by hand")
    CLForgetCache
    Set tampered = CLLoad(id, status)
    Check status = "modified", "A record edited outside Word is detected"
    Check CLGet(tampered, "title") = "Edited by hand", "It is still readable rather than lost"
    CLAcceptModified id
    Set tampered = CLLoad(id, status)
    Check status = "ok", "It can be accepted deliberately"

    Dim usage As Object
    CLTouchUsage id: CLTouchUsage id
    CLForgetCache
    Set usage = RowFor(id)
    Check Not usage Is Nothing, "An accepted record still appears in the library"
    If Not usage Is Nothing Then
        Check usage("usedCount") = "2", "Reuse is counted"
    End If
    Check CLRevision(CLLoad(id)) = CLRevision(tampered), "Counting reuse does not create a version of the wording"
End Sub

Private Function RowFor(ByVal id As String) As Object
    Dim r As Variant
    For Each r In CLList()
        If r("id") = id Then Set RowFor = r: Exit Function
    Next
End Function

' ---------- searching ----------

Private Sub CheckSearch()
    Note "Finding things"
    Dim e As Object, made As Variant, pair As Variant
    For Each pair In Array("Intellectual property licence|The Licensor grants a licence to the Background IP.|IP", _
                           "Principal contractor|The Principal shall procure that each participant complies.|Construction", _
                           "Force majeure|Neither party is liable for delay caused by a Force Majeure Event.|Risk", _
                           "Liability cap|The Supplier total liability shall not exceed the Charges paid.|Liability")
        made = Split(CStr(pair), "|")
        Set e = CLNew()
        CLSet e, "title", CStr(made(0)): CLSet e, "plain", CStr(made(1)): CLSet e, "topic", CStr(made(2))
        CLSet e, "fingerprint", CLFingerprint(CStr(made(1)))
        CLSave e, 0
    Next
    CLForgetCache
    Dim all As Collection: Set all = CLList()

    Check Hits(all, """force majeure""") = 1, "A phrase in quotation marks finds the clause, not nothing"
    Check Hits(all, "force majeure") = 1, "The same words without quotation marks also find it"
    Check Hits(all, "licence") = 1, "A single word narrows the list"
    Check Hits(all, "liability cap") = 1, "Two words narrow it further"
    Check Hits(all, "frustration") = 0, "Words that appear nowhere find nothing"
    Check Hits(all, "INTELLECTUAL") = 1, "Capital letters make no difference"
    Check Hits(all, "") >= 4, "An empty search shows everything active"

    ' Ranking is what stops short legal words being useless. "ip" appears
    ' inside Principal and participant too, but the right clause comes first.
    Dim ranked As Collection
    Set ranked = CLSearch(all, "ip", "All", "Best match")
    Check ranked.count >= 2, "A short word still finds everything that contains it"
    Check ranked(1)("title") = "Intellectual property licence", "The clause that is really about it is ranked first"

    Set ranked = CLSearch(all, "", "All", "Title")
    Check ranked(1)("title") < ranked(ranked.count)("title"), "Sorting by name works"

    Dim duplicate As Object
    Set duplicate = CLFindFingerprint(all, CLFingerprint("The   Licensor grants a licence to the Background IP."))
    Check Not duplicate Is Nothing, "The same wording captured twice is recognised despite different spacing"
    Check CLFindFingerprint(all, CLFingerprint("Entirely different wording.")) Is Nothing, "Different wording is not treated as a duplicate"
End Sub

Private Function Hits(ByVal all As Collection, ByVal query As String) As Long
    Hits = CLSearch(all, query, "All", "Best match").count
End Function

' ---------- history, backup, repair ----------

Private Sub CheckRecovery()
    Note "Previous versions, backup and repair"
    Dim e As Object, id As String, issues As String, versions As Collection
    Set e = CLNew()
    CLSet e, "title", "Payment terms": CLSet e, "plain", "Thirty days from invoice."
    CLSave e, 0, "<payload>v1</payload>", True
    id = CLEntryId(e)
    Set e = CLLoad(id): CLSet e, "plain", "Sixty days from invoice."
    CLSave e, 1, "<payload>v2</payload>", True

    Set versions = CLHistory(id, issues)
    Check versions.count = 1, "The replaced version is listed"
    Check versions(1)("plain") = "Thirty days from invoice.", "Its wording is the wording it had"
    Check CLPayloadForRevision(id, 1) = "<payload>v1</payload>", "Its formatting is the formatting it had"
    Check CLPayloadForRevision(id, 2) = "<payload>v2</payload>", "The current formatting resolves correctly too"

    CLRestoreRevision id, CStr(versions(1)("path")), 2, False
    Set e = CLLoad(id)
    Check CLRevision(e) = 3 And CLGet(e, "plain") = "Thirty days from invoice.", "Restoring makes a new version rather than rewriting history"
    Check CLHistory(id, issues).count = 2, "Both earlier versions are still there"

    ' A library with a damaged file must still be backed up. This is the
    ' failure that matters: you need a copy most when something is wrong.
    Dim broken As String
    broken = CLRoot() & "\entries\" & CLId() & ".xml"
    CLWrite broken, "<entry>this file is not valid</entr"
    CLForgetCache
    Dim countedBefore As Long: countedBefore = CLList().count
    Check CLReadIssueCount = 1, "A damaged file is reported"
    Check countedBefore > 0, "The healthy items are still listed"

    Dim report As String, backup As String
    backup = CLBackupTo(Environ$("TEMP"), True, report)
    Check Len(backup) > 0 And CLExists(backup & "\backup-complete.xml"), "The backup finished despite the damaged file"
    Check CLExists(backup & "\entries\" & CLEntryId(e) & ".xml"), "A healthy item is in the backup"
    Check CLExists(backup & "\entries\" & CLEntryId(e) & ".rich"), "Its formatting is in the backup too"
    Check CLFileSha(backup & "\entries\" & CLEntryId(e) & ".xml") = CLFileSha(CLEntryPath(CLEntryId(e))), "The copy is identical to the original"

    Dim restoredReport As String, restored As String
    restored = CLRestoreBackupTo(backup, Environ$("TEMP"), restoredReport)
    Check CLExists(restored & "\library.xml"), "The backup restores into a new folder"
    Check CLDom(CLRead(restored & "\library.xml")).documentElement.getAttribute("id") <> _
          CLDom(CLRead(CLRoot() & "\library.xml")).documentElement.getAttribute("id"), _
          "The restored copy has its own identity, so a stale copy cannot be mistaken for the live one"
    Check CLExists(CLRoot() & "\library.xml"), "The working library is untouched by a restore"

    ' Repair: make the current file unreadable, then put it back from history.
    CLWrite CLEntryPath(id), "<entry>broken"
    CLForgetCache
    Dim outcome As String: outcome = CLRepairEntry(id)
    Check InStr(outcome, "Restored from revision") > 0, "A damaged item can be put back from its previous version"
    Set e = CLLoad(id)
    Check CLGet(e, "title") = "Payment terms", "The repaired item reads correctly again"

    On Error Resume Next
    CLFso().DeleteFolder backup, True
    CLFso().DeleteFolder restored, True
    On Error GoTo 0
End Sub

' ---------- Word itself ----------

Private Sub CheckWord()
    Note "Capturing from Word and inserting into Word"
    Dim source As Document, dest As Document, xml As String, plain As String, before As String
    Dim inserted As Range, t As Table, blocked As Boolean

    Set source = Documents.Add(Visible:=False)
    source.Content.Text = "Bold clause" & vbCr & "Italic clause" & vbCr
    source.Range(0, 4).Font.Bold = True
    source.Range(12, 18).Font.Italic = True
    before = source.Content.Text
    xml = CLCapturePayload(source.Content, True, plain)
    Check source.Content.Text = before, "Capturing does not change the document you captured from"
    Set dest = Documents.Add(Visible:=False)
    dest.Content.InsertXML xml
    Check dest.Range(0, 4).Font.Bold = True, "Bold survives capture and insertion"
    Check dest.Range(12, 18).Font.Italic = True, "Italic survives capture and insertion"
    Check CLPlainWording(dest.Content.Text) = plain, "The inserted wording is the wording that was saved"
    Check InStr(xml, "sectPr") = 0, "Page setup is stripped, so a clause cannot change the contract's layout"
    dest.Close wdDoNotSaveChanges: source.Close wdDoNotSaveChanges

    Set source = Documents.Add(Visible:=False)
    source.Content.Text = "First provision" & vbCr & "Second provision" & vbCr
    source.Content.ListFormat.ApplyNumberDefault
    xml = CLCapturePayload(source.Content, True, plain)
    Check InStr(CLCapturePreview, "1.") > 0 Or InStr(CLCapturePreview, "1)") > 0, "The preview shows the numbering, which plain text leaves out"
    Set dest = Documents.Add(Visible:=False): dest.Content.InsertXML xml
    Check dest.Paragraphs(1).Range.ListFormat.ListType <> wdListNoNumbering, "Automatic numbering survives"
    Check dest.Paragraphs(2).Range.ListFormat.ListValue = 2, "The numbering carries on correctly"
    dest.Close wdDoNotSaveChanges: source.Close wdDoNotSaveChanges

    Set source = Documents.Add(Visible:=False)
    Set t = source.Tables.Add(source.Range(0, 0), 2, 2)
    t.Cell(1, 1).Range.Text = "Service": t.Cell(2, 2).Range.Text = "100"
    xml = CLCapturePayload(source.Content, True, plain)
    Set dest = Documents.Add(Visible:=False): dest.Content.InsertXML xml
    Check dest.Tables.count = 1 And dest.Tables(1).Range.Cells.count = 4, "A whole table survives with all its cells"
    dest.Close wdDoNotSaveChanges: source.Close wdDoNotSaveChanges

    Set source = Documents.Add(Visible:=False)
    source.Content.Text = "Old cap"
    source.TrackRevisions = True
    source.Range(0, 3).Text = "New"
    Dim revisionsBefore As Long: revisionsBefore = source.Revisions.count
    source.Comments.Add source.Range(0, 3), "Counterparty position - do not keep"
    xml = CLCapturePayload(source.Content, True, plain)
    Check source.Revisions.count = revisionsBefore And source.Comments.count = 1, "Your tracked changes and comments are left alone"
    Check InStr(plain, "New cap") > 0 And InStr(plain, "Old") = 0, "The final agreed wording is what gets saved"
    Set dest = Documents.Add(Visible:=False): dest.Content.InsertXML xml
    Check dest.Revisions.count = 0, "The saved wording carries no revision history into the next contract"
    Check dest.Comments.count = 0, "And no comments"
    dest.Close wdDoNotSaveChanges: source.Close wdDoNotSaveChanges

    ' Insertion behaviour in a live document.
    Set dest = Documents.Add(Visible:=False)
    dest.Content.Text = "Before" & vbCr & "After"
    dest.TrackRevisions = True
    before = dest.Content.Text
    CLInsertPayload dest.Range(7, 7), xml, True, inserted
    Check dest.TrackRevisions, "Inserting leaves Track Changes exactly as it was"
    Check dest.Revisions.count > 0, "With Track Changes on, the insertion is tracked"
    Check InStr(dest.Content.Text, "Before") > 0 And InStr(dest.Content.Text, "After") > 0, "Surrounding wording is untouched"
    dest.Undo 1
    Check dest.Content.Text = before, "One press of Undo puts the document back"

    dest.Content.InsertAfter " MYOWNEDIT"
    before = dest.Content.Text
    blocked = False
    On Error Resume Next
    CLInsertPayload dest.Range(0, 0), "this is not valid formatting", False, inserted
    blocked = (Err.number <> 0): Err.Clear
    On Error GoTo 0
    Check blocked And dest.Content.Text = before, "Damaged formatting is refused and the document is left alone"
    dest.Undo 1
    Check InStr(dest.Content.Text, "MYOWNEDIT") = 0, "A refused insertion does not use up your Undo - your own last edit is still what Undo reverses"

    dest.Protect wdAllowOnlyReading, NoReset:=True
    Check Len(CLInsertBlockedReason(dest.Content)) > 0, "A protected document is refused"
    dest.Unprotect
    dest.Close wdDoNotSaveChanges

    Set dest = Documents.Add(Visible:=False)
    dest.Tables.Add dest.Range(0, 0), 2, 2
    Dim reason As String: reason = CLInsertBlockedReason(dest.Range(1, 1))
    Check Len(reason) > 0, "Formatted wording is refused inside a table cell"
    Check InStr(reason, "Insert plain text") > 0, "And the refusal says what to do instead"
    dest.Close wdDoNotSaveChanges

    ' The whole capture path, with content that used to stop it dead. A
    ' footnote puts Chr(2) into Word's plain text, which XML cannot hold; before
    ' this was handled, capture failed with a message about library files.
    Set source = Documents.Add(Visible:=False)
    source.Content.Text = "The Supplier shall indemnify the Customer against all Losses."
    source.Footnotes.Add Range:=source.Range(12, 12), Text:="As defined in clause 1."
    Dim ok As Boolean, probe As Object
    On Error Resume Next
    Err.Clear
    xml = CLCapturePayload(source.Content, True, plain)
    Set probe = CLNew()
    CLSet probe, "title", "Footnote clause"
    CLSet probe, "plain", plain
    CLSave probe, 0, xml, True
    ok = (Err.number = 0): Err.Clear
    On Error GoTo 0
    Check ok, "Wording containing a footnote captures and saves end to end"
    If ok Then
        Check InStr(CLGet(CLLoad(CLEntryId(probe)), "plain"), "indemnify") > 0, "And reads back with its wording intact"
    End If
    source.Close wdDoNotSaveChanges
End Sub
