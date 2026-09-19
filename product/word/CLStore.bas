Attribute VB_Name = "CLStore"
' Clause Library - the repository. Word is the only writer.
'
' Layout of a library folder:
'   library.xml                  identity marker
'   entries\<id>.xml             metadata only, one to two kilobytes
'   entries\<id>.rich            the Word formatting payload, read only when needed
'   history\<id>-r<0000000001>.xml / .rich   every superseded revision
'   writer.lock                  serialises writers across Word processes
'
' Keeping the payload out of the metadata file is what lets searching stay
' instant: listing a thousand clauses reads about two megabytes, not eighty.
Option Explicit

Public CLTestRoot As String
Public CLReadIssues As String
Public CLReadIssueCount As Long
Private Const APPKEY As String = "ClauseLibraryPersonal"

Private rowCache As Object          ' path -> Array(signature, row)
Private cachedRoot As String

' ---------- the connected library ----------

Public Function CLRoot() As String
    If Len(CLTestRoot) > 0 Then CLRoot = CLTestRoot: Exit Function
    Dim connection As String, fields As Variant, marker As Object
    connection = GetSetting(APPKEY, "Library", "Connection", "")
    If Len(connection) = 0 Then CLFail "No library is connected yet. Open Start Here and choose Set up Clause Library."
    fields = Split(connection, "|")
    If UBound(fields) <> 1 Then CLFail "The saved library location could not be read. Use Locate a library to reconnect your folder."
    CLRoot = fields(1)
    If Not CLExists(CLRoot & "\library.xml") Then
        CLFail "Your library folder is not available at:" & vbCrLf & CLRoot & vbCrLf & vbCrLf & _
               "Nothing was deleted and no empty replacement was created. If the folder moved, or is on a drive that is not connected, use Locate a library."
    End If
    Set marker = CLMarker(CLRoot)
    If marker.documentElement.getAttribute("id") <> fields(0) Then
        CLFail "The folder at " & CLRoot & " is not the library this copy of Word is connected to. Use Locate a library to choose deliberately."
    End If
End Function

Private Function CLMarker(ByVal root As String) As Object
    Dim d As Object: Set d = CLDom(CLRead(root & "\library.xml"))
    If d.documentElement.nodeName <> "library" Or d.documentElement.getAttribute("product") <> CL_PRODUCT Then
        CLFail "That folder is not a Clause Library."
    End If
    Dim schema As String: schema = d.documentElement.getAttribute("schema")
    If schema <> "1" And schema <> CL_SCHEMA Then
        CLFail "This library was made by a newer version of Clause Library (format " & schema & "). Your files were not changed. Install the newer version to open it."
    End If
    If Not CLIsId(CStr(d.documentElement.getAttribute("id"))) Then CLFail "This library's identity marker is damaged. Your files were preserved; restore a backup or use Locate a library."
    If Not CLFolderExists(root & "\entries") Or Not CLFolderExists(root & "\history") Then
        CLFail "This library folder is incomplete - its entries or history folder is missing. Restore a backup, or use Locate a library to find your original folder."
    End If
    Set CLMarker = d
End Function

Public Function CLLibrarySchema(ByVal root As String) As String
    CLLibrarySchema = CLMarker(root).documentElement.getAttribute("schema")
End Function

Public Sub CLCreateLibrary(ByVal root As String)
    If Len(root) > 160 Then CLFail "Choose a library folder closer to the top of the drive. This path is too long for Word to save into reliably."
    CLFolder root
    If CLExists(root & "\library.xml") Then CLFail "A library already exists in that folder. Use Locate a library to connect to it instead."
    CLFolder root & "\entries": CLFolder root & "\history"
    CLWrite root & "\library.xml", "<library product=""" & CL_PRODUCT & """ schema=""" & CL_SCHEMA & _
        """ id=""" & CLId() & """ created=""" & CLStamp() & """ madeBy=""" & CL_VERSION & """/>"
End Sub

Public Sub CLRemember(ByVal root As String)
    If Len(root) > 160 Then CLFail "Choose a library folder closer to the top of the drive. This path is too long for Word to save into reliably."
    Dim d As Object: Set d = CLMarker(root)
    If Not CLCloseManager() Then CLFail "The library window is still open with unsaved changes, so the library was not switched. Your current library is still connected."
    SaveSetting APPKEY, "Library", "Connection", d.documentElement.getAttribute("id") & "|" & root
    CLForgetCache
End Sub

Public Function CLSetting(ByVal name As String, ByVal fallback As String) As String
    CLSetting = GetSetting(APPKEY, "Options", name, fallback)
End Function

Public Sub CLSetSetting(ByVal name As String, ByVal value As String)
    SaveSetting APPKEY, "Options", name, value
End Sub

Public Sub CLForgetCache()
    Set rowCache = Nothing: cachedRoot = ""
End Sub

' ---------- moving a version 1 library to version 2 ----------

' Version 1 nested the whole Word payload inside the metadata file. Migration
' lifts it out to a sibling file. A complete backup is taken first and nothing
' is deleted, so this is reversible by copying the backup back.
Public Function CLNeedsMigration() As Boolean
    On Error Resume Next
    CLNeedsMigration = (CLLibrarySchema(CLRoot()) = "1")
End Function

Public Function CLMigrate(ByVal root As String, ByRef report As String) As Long
    Dim f As Object, d As Object, rich As String, moved As Long, seen As Long, folder As Variant
    Dim gate As Integer
    On Error GoTo Failed
    gate = FreeFile: Open root & "\writer.lock" For Binary Access Read Write Lock Read Write As #gate
    For Each folder In Array("entries", "history")
        For Each f In CLFso().GetFolder(root & "\" & folder).Files
            If LCase$(Right$(f.Name, 4)) = ".xml" Then
                seen = seen + 1
                Set d = CLDom(CLRead(f.path))
                If d.documentElement.nodeName = "entry" Then
                    rich = CLGet(d, "rich")
                    If Len(rich) > 0 Then
                        CLWrite Left$(f.path, Len(f.path) - 4) & ".rich", rich
                        CLRemove d, "rich"
                        CLSet d, "richHash", CLFileSha(Left$(f.path, Len(f.path) - 4) & ".rich")
                        moved = moved + 1
                    End If
                    If Len(CLGet(d, "fingerprint")) = 0 Then CLSet d, "fingerprint", CLFingerprint(CLGet(d, "plain"))
                    d.documentElement.setAttribute "schema", CL_SCHEMA
                    d.documentElement.setAttribute "digest", CLDigest(d)
                    CLAtomicWrite f.path, d.XML
                End If
            End If
        Next
    Next
    Dim marker As Object: Set marker = CLDom(CLRead(root & "\library.xml"))
    marker.documentElement.setAttribute "schema", CL_SCHEMA
    marker.documentElement.setAttribute "migrated", CLStamp()
    marker.documentElement.setAttribute "madeBy", CL_VERSION
    CLAtomicWrite root & "\library.xml", marker.XML
    Close #gate
    CLForgetCache
    report = seen & " records checked, " & moved & " Word payloads moved into their own files."
    CLMigrate = moved
    Exit Function
Failed:
    Dim message As String: message = CLExplain(Err.number, Err.Description)
    On Error Resume Next
    Close #gate
    On Error GoTo 0
    CLFail "The library was not fully updated. " & message & " Nothing was deleted; restore the backup taken a moment ago if you want to go back."
End Function

' ---------- records ----------

Public Function CLNew(Optional ByVal kind As String = "Clause") As Object
    Dim d As Object
    Set d = CLDom("<entry product=""" & CL_PRODUCT & """ schema=""" & CL_SCHEMA & """ revision=""0""><data/></entry>")
    d.documentElement.setAttribute "id", CLId()
    CLSet d, "kind", kind
    CLSet d, "created", CLStamp(): CLSet d, "state", "Active"
    CLSet d, "organised", "0": CLSet d, "favourite", "0"
    Set CLNew = d
End Function

Public Function CLGet(ByVal entry As Object, ByVal key As String) As String
    Dim n As Object: Set n = entry.selectSingleNode("/entry/data/" & key)
    If Not n Is Nothing Then CLGet = n.Text
End Function

Public Sub CLSet(ByVal entry As Object, ByVal key As String, ByVal value As String)
    Dim n As Object: Set n = entry.selectSingleNode("/entry/data/" & key)
    If n Is Nothing Then
        Set n = entry.createElement(key): entry.selectSingleNode("/entry/data").appendChild n
    End If
    n.Text = CLSafeText(value)
End Sub

Public Sub CLRemove(ByVal entry As Object, ByVal key As String)
    Dim n As Object: Set n = entry.selectSingleNode("/entry/data/" & key)
    If Not n Is Nothing Then n.parentNode.removeChild n
End Sub

Public Function CLEntryId(ByVal entry As Object) As String
    CLEntryId = entry.documentElement.getAttribute("id")
End Function

Public Function CLRevision(ByVal entry As Object) As Long
    CLRevision = CLng(Val(entry.documentElement.getAttribute("revision")))
End Function

Public Function CLDigest(ByVal entry As Object) As String
    CLDigest = CLSha(CLEntryId(entry) & vbLf & CStr(CLRevision(entry)) & vbLf & entry.selectSingleNode("/entry/data").XML)
End Function

' Identifies the same wording captured twice, ignoring spacing and case.
Public Function CLFingerprint(ByVal plain As String) As String
    Dim t As String: t = LCase$(CLSafeText(plain))
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.Pattern = "[^a-z0-9]+"
    t = Trim$(re.Replace(t, " "))
    If Len(t) = 0 Then Exit Function
    CLFingerprint = CLSha(t)
End Function

Public Function CLEntryPath(ByVal id As String) As String
    CLEntryPath = CLRoot() & "\entries\" & id & ".xml"
End Function

Public Function CLPayloadPath(ByVal id As String) As String
    CLPayloadPath = CLRoot() & "\entries\" & id & ".rich"
End Function

' ---------- reading ----------

' Returns the entry. On a digest mismatch the record is still returned, with
' status "modified", rather than being hidden: a file someone edited by hand
' should be visible and repairable, not lost.
Public Function CLLoadPath(ByVal path As String, Optional ByRef status As String) As Object
    Dim d As Object: Set d = CLDom(CLRead(path))
    status = "ok"
    If d.documentElement.nodeName <> "entry" Then CLFail "This file is not a Clause Library entry."
    If d.documentElement.getAttribute("product") <> CL_PRODUCT Then CLFail "This entry belongs to a different application."
    Dim schema As String: schema = d.documentElement.getAttribute("schema")
    If schema <> "1" And schema <> CL_SCHEMA Then CLFail "This entry was written by a newer version of Clause Library."
    If Not CLIsId(CLEntryId(d)) Then CLFail "This entry's identifier is not valid."
    If CLRevision(d) < 1 Then CLFail "This entry has no saved revision."
    If d.documentElement.getAttribute("digest") <> CLDigest(d) Then status = "modified"
    Set CLLoadPath = d
End Function

Public Function CLLoad(ByVal id As String, Optional ByRef status As String) As Object
    If Not CLIsId(id) Then CLFail "That is not a valid item identifier."
    Dim path As String: path = CLEntryPath(id)
    If Not CLExists(path) Then CLFail "That item is no longer in your library. It may have been removed outside Word; check Previous versions or restore a backup."
    Dim e As Object: Set e = CLLoadPath(path, status)
    If CLEntryId(e) <> id Then CLFail "This entry's file name and identifier disagree. The file was preserved and not changed."
    Set CLLoad = e
End Function

Public Function CLReadPayload(ByVal entry As Object) As String
    Dim path As String, expected As String
    If Len(CLGet(entry, "rich")) > 0 Then CLReadPayload = CLGet(entry, "rich"): Exit Function   ' version 1 record
    path = CLRoot() & "\entries\" & CLEntryId(entry) & ".rich"
    If Not CLExists(path) Then Exit Function
    expected = CLGet(entry, "richHash")
    If Len(expected) > 0 Then
        If CLFileSha(path) <> expected Then
            CLFail "The saved Word formatting for this item does not match its record, so it was not used. Use Insert plain text, or restore an earlier version from Previous versions."
        End If
    End If
    CLReadPayload = CLRead(path)
End Function

' ---------- writing ----------

' Every change to a record goes through here. In order: take the cross-process
' lock, refuse a stale editor, keep the outgoing revision in history, write the
' payload, then swap the metadata file in atomically.
Public Sub CLSave(ByVal entry As Object, ByVal expectedRevision As Long, Optional ByVal payload As String, Optional ByVal payloadChanged As Boolean = False)
    Dim root As String, file As String, lockFile As Integer, old As Object
    Dim current As Long, backup As String, id As String, payloadFile As String
    On Error GoTo Failed
    id = CLEntryId(entry)
    If Not CLIsId(id) Then CLFail "That is not a valid item identifier."
    root = CLRoot(): file = root & "\entries\" & id & ".xml": payloadFile = root & "\entries\" & id & ".rich"
    lockFile = FreeFile
    Open root & "\writer.lock" For Binary Access Read Write Lock Read Write As #lockFile
    If CLExists(file) Then
        Set old = CLLoadPath(file)
        current = CLRevision(old)
        If current <> expectedRevision Then
            CLFail "This item was changed in another window since you opened it. Nothing was overwritten. Close this item, reopen it, and make your change again."
        End If
        backup = root & "\history\" & id & "-r" & Format$(current, "0000000000")
        If Not CLExists(backup & ".xml") Then CLWrite backup & ".xml", old.XML
        ' The payload is only copied into history when it is about to be
        ' replaced. Earlier revisions that did not change it resolve forward
        ' to the first later snapshot that did - see CLPayloadForRevision.
        If payloadChanged And CLExists(payloadFile) And Not CLExists(backup & ".rich") Then
            CLCopyFile payloadFile, backup & ".rich"
        End If
    ElseIf expectedRevision <> 0 Then
        CLFail "The saved item is missing from your library, so your copy was not written over it. Check Previous versions or restore a backup."
    End If
    If payloadChanged Then
        If Len(payload) > 0 Then
            CLAtomicWrite payloadFile, payload
            CLSet entry, "richHash", CLFileSha(payloadFile)
            CLSet entry, "richBytes", CStr(Len(payload))
        Else
            If CLExists(payloadFile) Then CLFso().DeleteFile payloadFile
            CLRemove entry, "richHash": CLRemove entry, "richBytes"
        End If
    End If
    entry.documentElement.setAttribute "revision", CStr(current + 1)
    entry.documentElement.setAttribute "schema", CL_SCHEMA
    CLSet entry, "updated", CLStamp()
    entry.documentElement.setAttribute "digest", CLDigest(entry)
    ' Confirm the exact text we are about to commit still reads back as this record.
    Dim verify As Object: Set verify = CLDom(entry.XML)
    If CLDigest(verify) <> entry.documentElement.getAttribute("digest") Then
        CLFail "This record could not be written reliably and was not saved. Nothing was changed."
    End If
    CLAtomicWrite file, entry.XML
    Close #lockFile: lockFile = 0
    CLInvalidate file
    Exit Sub
Failed:
    Dim message As String: message = CLExplain(Err.number, Err.Description)
    On Error Resume Next
    If lockFile <> 0 Then Close #lockFile
    On Error GoTo 0
    CLFail message
End Sub

' Explicitly accept a record that someone edited outside Word.
Public Sub CLAcceptModified(ByVal id As String)
    Dim e As Object, status As String
    Set e = CLLoad(id, status)
    If status = "ok" Then Exit Sub
    e.documentElement.setAttribute "digest", CLDigest(e)
    Dim lockFile As Integer: lockFile = FreeFile
    Open CLRoot() & "\writer.lock" For Binary Access Read Write Lock Read Write As #lockFile
    CLAtomicWrite CLEntryPath(id), e.XML
    Close #lockFile
    CLInvalidate CLEntryPath(id)
End Sub

' How often wording gets used is the best guide to what is worth reaching for,
' but it is not part of the wording. It lives in its own small file so that
' reusing a clause does not manufacture a revision of it.
Public Function CLUsagePath(ByVal id As String) As String
    CLUsagePath = CLRoot() & "\entries\" & id & ".use"
End Function

Public Sub CLTouchUsage(ByVal id As String)
    On Error Resume Next
    Dim path As String, count As Long, d As Object
    path = CLUsagePath(id)
    If CLExists(path) Then
        Set d = CLDom(CLRead(path))
        count = Val(d.documentElement.getAttribute("count"))
    End If
    CLWrite path, "<use count=""" & (count + 1) & """ last=""" & CLStamp() & """/>"
    CLInvalidate CLEntryPath(id)
End Sub

Public Sub CLReadUsage(ByVal row As Object)
    On Error Resume Next
    Dim path As String, d As Object
    row("usedCount") = "0": row("lastUsed") = ""
    path = CLRoot() & "\entries\" & row("id") & ".use"
    If Not CLExists(path) Then Exit Sub
    Set d = CLDom(CLRead(path))
    row("usedCount") = CStr(Val(d.documentElement.getAttribute("count")))
    row("lastUsed") = CStr(d.documentElement.getAttribute("last"))
End Sub

' The payload as it stood at a given revision.
Public Function CLPayloadForRevision(ByVal id As String, ByVal revision As Long) As String
    Dim root As String, candidate As String, r As Long, newest As Long
    root = CLRoot()
    newest = revision + 400
    For r = revision To newest
        candidate = root & "\history\" & id & "-r" & Format$(r, "0000000000") & ".rich"
        If CLExists(candidate) Then CLPayloadForRevision = CLRead(candidate): Exit Function
        If Not CLExists(root & "\history\" & id & "-r" & Format$(r, "0000000000") & ".xml") Then Exit For
    Next
    candidate = root & "\entries\" & id & ".rich"
    If CLExists(candidate) Then CLPayloadForRevision = CLRead(candidate)
End Function

' ---------- listing ----------

Private Sub CLInvalidate(ByVal path As String)
    If rowCache Is Nothing Then Exit Sub
    If rowCache.Exists(path) Then rowCache.Remove path
End Sub

' One dictionary per entry, with the searchable text worked out once. Files
' that have not changed since the last listing are reused untouched, so a
' refresh after ticking Favourite costs almost nothing.
Public Function CLList() As Collection
    Dim result As New Collection, f As Object, root As String, signature As String
    Dim row As Object, status As String, e As Object
    root = CLRoot()
    CLReadIssues = "": CLReadIssueCount = 0
    If rowCache Is Nothing Or cachedRoot <> root Then Set rowCache = CreateObject("Scripting.Dictionary"): cachedRoot = root
    For Each f In CLFso().GetFolder(root & "\entries").Files
        If LCase$(Right$(f.Name, 4)) = ".xml" Then
            signature = CStr(f.Size) & "|" & CStr(CDbl(f.DateLastModified))
            If rowCache.Exists(f.path) Then
                If rowCache(f.path)(0) = signature Then result.Add rowCache(f.path)(1): GoTo NextFile
                rowCache.Remove f.path
            End If
            Set e = Nothing: status = ""
            On Error Resume Next
            Set e = CLLoadPath(f.path, status)
            If Not e Is Nothing Then
                If f.Name <> CLEntryId(e) & ".xml" Then
                    Set e = Nothing
                    Err.Raise vbObjectError + CL_ERR, , "The file name and the identifier inside it do not match."
                End If
            End If
            If Err.number <> 0 Then
                CLReadIssues = CLReadIssues & "  " & f.Name & " - " & Err.Description & vbCrLf
                CLReadIssueCount = CLReadIssueCount + 1
                Err.Clear
            End If
            On Error GoTo 0
            If Not e Is Nothing Then
                Set row = CLMakeRow(e, status)
                rowCache.Add f.path, Array(signature, row)
                result.Add row
            End If
        End If
NextFile:
    Next
    Set CLList = result
End Function

Private Function CLMakeRow(ByVal e As Object, ByVal status As String) As Object
    Dim r As Object: Set r = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In Array("title", "plain", "preview", "topic", "tags", "family", "role", "applicability", _
                        "notes", "comment", "kind", "state", "organised", "favourite", _
                        "created", "updated", "source", "fingerprint", "richHash")
        r(CStr(k)) = CLGet(e, CStr(k))
    Next
    r("id") = CLEntryId(e)
    CLReadUsage r
    r("revision") = CLRevision(e)
    r("status") = status
    r("hasRich") = (Len(r("richHash")) > 0) Or (Len(CLGet(e, "rich")) > 0)
    r("sTitle") = CLNormalise(r("title"))
    r("sMeta") = CLNormalise(r("topic") & " " & r("tags") & " " & r("family") & " " & r("role") & " " & r("kind"))
    r("sGuide") = CLNormalise(r("applicability") & " " & r("notes"))
    r("sBody") = CLNormalise(r("plain") & " " & r("comment"))
    r("sSource") = CLNormalise(r("source"))
    Set CLMakeRow = r
End Function

' Lower case, punctuation folded to spaces, wrapped in spaces so that a whole
' word can be found with an ordinary InStr.
Public Function CLNormalise(ByVal value As String) As String
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True: re.Pattern = "[^a-z0-9]+"
    End If
    CLNormalise = " " & Trim$(re.Replace(LCase$(CLSafeText(value)), " ")) & " "
End Function

' ---------- searching ----------

' Splits a query into terms, honouring "quoted phrases". A lawyer who types
' quotation marks gets a phrase search rather than nothing at all.
Public Function CLTerms(ByVal query As String) As Collection
    Dim result As New Collection, i As Long, ch As String, buffer As String, inQuote As Boolean
    query = Trim$(query)
    For i = 1 To Len(query)
        ch = Mid$(query, i, 1)
        If ch = Chr$(34) Or ch = Chr$(147) Or ch = Chr$(148) Then
            If inQuote And Len(Trim$(buffer)) > 0 Then result.Add CLNormalise(buffer)
            If inQuote Then buffer = ""
            inQuote = Not inQuote
        ElseIf ch = " " And Not inQuote Then
            If Len(Trim$(buffer)) > 0 Then result.Add CLNormalise(buffer)
            buffer = ""
        Else
            buffer = buffer & ch
        End If
    Next
    If Len(Trim$(buffer)) > 0 Then result.Add CLNormalise(buffer)
    Set CLTerms = result
End Function

Private Function CLFieldScore(ByVal haystack As String, ByVal needle As String) As Long
    ' needle arrives normalised and space-wrapped, e.g. " indemnity "
    If Len(haystack) <= 2 Then Exit Function
    If InStr(haystack, needle) > 0 Then CLFieldScore = 3: Exit Function          ' whole word or phrase
    If InStr(haystack, Left$(needle, Len(needle) - 1)) > 0 Then CLFieldScore = 2: Exit Function  ' word beginning
    If InStr(haystack, Mid$(needle, 2, Len(needle) - 2)) > 0 Then CLFieldScore = 1               ' inside a word
End Function

' Every term must appear somewhere (a narrowing search, as people expect), and
' where it appears decides the ranking.
Public Function CLScore(ByVal row As Object, ByVal terms As Collection) As Long
    Dim t As Variant, best As Long, total As Long, s As Long
    If terms.count = 0 Then CLScore = 1: Exit Function
    For Each t In terms
        best = 0
        s = CLFieldScore(row("sTitle"), CStr(t)): If s * 6 > best Then best = s * 6
        s = CLFieldScore(row("sMeta"), CStr(t)): If s * 4 > best Then best = s * 4
        s = CLFieldScore(row("sGuide"), CStr(t)): If s * 2 > best Then best = s * 2
        s = CLFieldScore(row("sBody"), CStr(t)): If s * 1 > best Then best = s
        s = CLFieldScore(row("sSource"), CStr(t)): If s * 1 > best Then best = s
        If best = 0 Then Exit Function
        total = total + best
    Next
    If row("favourite") = "1" Then total = total + 2
    CLScore = total
End Function

Public Function CLInView(ByVal row As Object, ByVal view As String) As Boolean
    Select Case view
        Case "Archived": CLInView = (row("state") = "Archived")
        Case "Trash": CLInView = (row("state") = "Trash")
        Case "Needs attention": CLInView = (row("state") = "Active" And row("status") <> "ok")
        Case "Not organised": CLInView = (row("state") = "Active" And row("organised") <> "1")
        Case "Favourites": CLInView = (row("state") = "Active" And row("favourite") = "1")
        Case Else: CLInView = (row("state") = "Active")
    End Select
End Function

' Returns rows in display order. sortBy: Best match, Recently used, Recently added, Title.
Public Function CLSearch(ByVal rows As Collection, ByVal query As String, ByVal view As String, ByVal sortBy As String) As Collection
    Dim terms As Collection, r As Variant, keep As Collection, keys() As String, items() As Object
    Dim n As Long, i As Long, score As Long
    Set terms = CLTerms(query)
    Set keep = New Collection
    For Each r In rows
        If CLInView(r, view) Then
            score = CLScore(r, terms)
            If score > 0 Then r("score") = score: keep.Add r
        End If
    Next
    Set CLSearch = keep
    If keep.count < 2 Then Exit Function
    ReDim keys(1 To keep.count): ReDim items(1 To keep.count)
    For i = 1 To keep.count
        Set items(i) = keep(i)
        Select Case sortBy
            Case "Recently used": keys(i) = CLPad(items(i)("lastUsed"), 19) & CLPad(items(i)("updated"), 19)
            Case "Recently added": keys(i) = CLPad(items(i)("created"), 19)
            Case "Title": keys(i) = LCase$(items(i)("title"))
            Case Else: keys(i) = Format$(9999 - CLCap(items(i)("score")), "0000") & CLPad(items(i)("lastUsed"), 19)
        End Select
    Next
    CLSortPairs keys, items
    Dim ordered As New Collection
    If sortBy = "Title" Or sortBy = "Best match" Then
        For i = 1 To UBound(items): ordered.Add items(i): Next
    Else
        For i = UBound(items) To 1 Step -1: ordered.Add items(i): Next
    End If
    Set CLSearch = ordered
End Function

Private Function CLCap(ByVal value As Long) As Long
    CLCap = value: If CLCap > 9998 Then CLCap = 9998
    If CLCap < 0 Then CLCap = 0
End Function

Private Function CLPad(ByVal value As String, ByVal width As Long) As String
    CLPad = Left$(value & String$(width, " "), width)
End Function

' Straightforward insertion sort. Libraries of this size never make it matter,
' and it is easy to read three years from now.
Private Sub CLSortPairs(ByRef keys() As String, ByRef items() As Object)
    Dim i As Long, j As Long, k As String, it As Object
    For i = LBound(keys) + 1 To UBound(keys)
        k = keys(i): Set it = items(i): j = i - 1
        Do While j >= LBound(keys)
            If keys(j) <= k Then Exit Do
            keys(j + 1) = keys(j): Set items(j + 1) = items(j): j = j - 1
        Loop
        keys(j + 1) = k: Set items(j + 1) = it
    Next
End Sub

Public Function CLFindFingerprint(ByVal rows As Collection, ByVal fingerprint As String) As Object
    Dim r As Variant
    If Len(fingerprint) = 0 Then Exit Function
    For Each r In rows
        If r("fingerprint") = fingerprint And r("state") = "Active" Then Set CLFindFingerprint = r: Exit Function
    Next
End Function
