Attribute VB_Name = "CLRecovery"
' Clause Library - previous versions, comparison, backup, restore and repair.
'
' The rule this module is built around: the moment your library develops a
' problem is the moment you most need a copy of it. Nothing here refuses to
' work because one file is damaged.
Option Explicit

' ---------- previous versions ----------

' Newest first. A damaged history file is reported, never fatal.
Public Function CLHistory(ByVal id As String, Optional ByRef issues As String) As Collection
    Dim result As New Collection, f As Object, e As Object, rows() As Object, keys() As String
    Dim n As Long, i As Long, row As Object, prefix As String
    prefix = id & "-r"
    For Each f In CLFso().GetFolder(CLRoot() & "\history").Files
        If LCase$(Right$(f.Name, 4)) = ".xml" Then
            If Left$(f.Name, Len(prefix)) = prefix Then
                Set e = Nothing
                On Error Resume Next
                Set e = CLLoadPath(f.path)
                If Err.number <> 0 Then issues = issues & "  " & f.Name & vbCrLf: Err.Clear
                On Error GoTo 0
                If Not e Is Nothing Then
                    If CLEntryId(e) = id Then
                        Set row = CreateObject("Scripting.Dictionary")
                        row("path") = f.path
                        row("revision") = CLRevision(e)
                        row("updated") = CLGet(e, "updated")
                        row("title") = CLGet(e, "title")
                        row("plain") = CLGet(e, "plain")
                        row("notes") = CLGet(e, "notes")
                        row("state") = CLGet(e, "state")
                        result.Add row
                    End If
                End If
            End If
        End If
    Next
    If result.count < 2 Then Set CLHistory = result: Exit Function
    ReDim rows(1 To result.count): ReDim keys(1 To result.count)
    For i = 1 To result.count
        Set rows(i) = result(i): keys(i) = Format$(rows(i)("revision"), "0000000000")
    Next
    Dim k As String, it As Object, j As Long
    For i = 2 To UBound(keys)
        k = keys(i): Set it = rows(i): j = i - 1
        Do While j >= 1
            If keys(j) <= k Then Exit Do
            keys(j + 1) = keys(j): Set rows(j + 1) = rows(j): j = j - 1
        Loop
        keys(j + 1) = k: Set rows(j + 1) = it
    Next
    Dim ordered As New Collection
    For i = UBound(rows) To 1 Step -1: ordered.Add rows(i): Next
    Set CLHistory = ordered
End Function

Public Sub CLRestoreRevision(ByVal id As String, ByVal snapshot As String, ByVal expectedRevision As Long, ByVal wordingOnly As Boolean)
    Dim old As Object, current As Object, payload As String
    Set old = CLLoadPath(snapshot)
    If CLEntryId(old) <> id Then CLFail "That previous version belongs to a different item and was not restored."
    payload = CLPayloadForRevision(id, CLRevision(old))
    If wordingOnly Then
        Set current = CLLoad(id)
        CLSet current, "plain", CLGet(old, "plain")
        CLSet current, "comment", CLGet(old, "comment")
        CLSet current, "fingerprint", CLFingerprint(CLGet(old, "plain"))
        CLSave current, expectedRevision, payload, True
    Else
        CLSave old, expectedRevision, payload, True
    End If
End Sub

' Word's own comparison, which is the tool a lawyer already trusts for this.
Public Function CLCompareRevision(ByVal id As String, ByVal snapshot As String) As Document
    Dim old As Object, current As Object, a As Document, b As Document, result As Document
    Set old = CLLoadPath(snapshot)
    Set current = CLLoad(id)
    Set a = CLDraftOf(old, "Revision " & CLRevision(old))
    Set b = CLDraftOf(current, "Current (revision " & CLRevision(current) & ")")
    Set result = Application.CompareDocuments(OriginalDocument:=a, RevisedDocument:=b, _
        Destination:=wdCompareDestinationNew, Granularity:=wdGranularityWordLevel, _
        CompareFormatting:=True, CompareCaseChanges:=True, CompareWhitespace:=False, _
        CompareTables:=True, CompareHeaders:=False, CompareFootnotes:=False, _
        CompareTextboxes:=False, CompareFields:=False, CompareComments:=False, _
        CompareMoves:=True, RevisedAuthor:="Current version", IgnoreAllComparisonWarnings:=True)
    a.Close wdDoNotSaveChanges: b.Close wdDoNotSaveChanges
    On Error Resume Next
    result.Windows(1).Caption = "Clause Library comparison - " & Left$(CLGet(current, "title"), 60)
    result.Saved = True
    On Error GoTo 0
    Set CLCompareRevision = result
End Function

Private Function CLDraftOf(ByVal entry As Object, ByVal label As String) As Document
    Dim d As Document, payload As String
    Set d = Documents.Add(Visible:=False)
    d.TrackRevisions = False
    payload = CLGet(entry, "rich")
    If Len(payload) = 0 Then payload = CLPayloadForRevision(CLEntryId(entry), CLRevision(entry))
    On Error Resume Next
    If Len(payload) > 0 Then d.Content.InsertXML payload
    On Error GoTo 0
    If Len(CLSafeText(d.Content.Text)) < 2 Then d.Content.Text = Replace$(CLGet(entry, "plain"), vbLf, vbCr)
    Set CLDraftOf = d
End Function

Public Sub CLShowHistory(ByVal id As String)
    On Error GoTo Failed
    Dim f As Object: Set f = VBA.UserForms.Add("frmHistory")
    f.LoadHistory id
    f.Show
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbExclamation, "Previous versions are not available"
End Sub

' ---------- repair ----------

' Puts an entry back from its most recent readable previous version. This is
' the way out when a file has been edited or damaged outside Word.
Public Function CLRepairEntry(ByVal id As String) As String
    Dim versions As Collection, v As Variant, issues As String
    Set versions = CLHistory(id, issues)
    If versions.count = 0 Then
        CLRepairEntry = "There is no previous version of this item to restore from. Its file is still in the entries folder, unchanged."
        Exit Function
    End If
    Dim newest As Object: Set newest = versions(1)
    Dim current As Long
    On Error Resume Next
    current = CLRevision(CLLoad(id))
    On Error GoTo 0
    Dim e As Object: Set e = CLLoadPath(newest("path"))
    Dim damaged As String: damaged = CLEntryPath(id) & ".damaged-" & Format$(Now, "yyyymmdd-hhnnss")
    If CLExists(CLEntryPath(id)) Then CLFso().MoveFile CLEntryPath(id), damaged
    CLSave e, 0, CLPayloadForRevision(id, newest("revision")), True
    CLRepairEntry = "Restored from revision " & newest("revision") & " (" & CLNiceDate(newest("updated")) & ")." & vbCrLf & _
                    "The unreadable file was kept, renamed to:" & vbCrLf & damaged
End Function

' ---------- backup ----------

' Copies everything, verifies what it can, and never stops because one file is
' damaged: a damaged file is copied as-is and reported.
Public Function CLBackupTo(ByVal parent As String, Optional ByVal quiet As Boolean = False, Optional ByRef report As String) As String
    Dim source As String, destination As String, gate As Integer, folder As Variant, f As Object
    Dim copied As Long, unverified As Long, problems As String, message As String
    On Error GoTo Failed
    source = CLRoot()
    destination = parent & "\Clause Library backup " & Format$(Now, "yyyy-mm-dd hhnnss")
    If CLFolderExists(destination) Then destination = destination & "-" & Left$(CLId(), 4)
    gate = FreeFile
    Open source & "\writer.lock" For Binary Access Read Write Lock Read Write As #gate
    CLFolder destination: CLFolder destination & "\entries": CLFolder destination & "\history"
    For Each folder In Array("entries", "history")
        For Each f In CLFso().GetFolder(source & "\" & folder).Files
            If CLBackupWanted(f.Name) Then
                On Error Resume Next
                Err.Clear
                CLFso().CopyFile f.path, destination & "\" & folder & "\" & f.Name, True
                If Err.number <> 0 Then
                    problems = problems & "  " & f.Name & " - could not be copied" & vbCrLf
                    Err.Clear
                Else
                    copied = copied + 1
                    If CLFileSha(f.path) <> CLFileSha(destination & "\" & folder & "\" & f.Name) Then
                        unverified = unverified + 1
                        problems = problems & "  " & f.Name & " - copied but could not be verified" & vbCrLf
                    End If
                    Err.Clear
                End If
                On Error GoTo Failed
            End If
        Next
    Next
    CLFso().CopyFile source & "\library.xml", destination & "\library.xml", True
    ' The completion marker is written last, so an interrupted backup can never
    ' be mistaken for a finished one.
    CLWrite destination & "\backup-complete.xml", "<backup product=""" & CL_PRODUCT & """ schema=""" & CL_SCHEMA & _
        """ files=""" & copied & """ unverified=""" & unverified & """ created=""" & CLStamp() & """ madeBy=""" & CL_VERSION & """/>"
    Close #gate: gate = 0
    CLSetSetting "LastBackup", CLStamp() & " to " & destination
    report = copied & " files copied."
    If Len(problems) > 0 Then report = report & vbCrLf & vbCrLf & "These files need your attention:" & vbCrLf & problems
    CLBackupTo = destination
    Exit Function
Failed:
    message = CLExplain(Err.number, Err.Description)
    On Error Resume Next
    If gate <> 0 Then Close #gate
    On Error GoTo 0
    CLFail "The backup did not finish. " & message & " Your working library was not changed."
End Function

Private Function CLBackupWanted(ByVal name As String) As Boolean
    Dim dot As Long: dot = InStrRev(name, ".")
    If dot = 0 Then Exit Function
    Dim ext As String: ext = LCase$(Mid$(name, dot))
    CLBackupWanted = (ext = ".xml" Or ext = ".rich" Or ext = ".use")
End Function

Public Sub CLBackUp()
    On Error GoTo Failed
    If Not CLEnsureReady() Then Exit Sub
    Dim picker As Object: Set picker = Application.FileDialog(4)
    picker.Title = "Choose where to keep a backup of your library"
    If picker.Show <> -1 Then Exit Sub
    CLRequireLocalPath CStr(picker.SelectedItems(1)), "The folder you chose"
    If Not CLConfirmSyncedFolder(CStr(picker.SelectedItems(1)), "Your whole library, including its private notes,") Then Exit Sub
    Dim report As String, destination As String
    destination = CLBackupTo(picker.SelectedItems(1), False, report)
    MsgBox "Backup saved to:" & vbCrLf & destination & vbCrLf & vbCrLf & report & vbCrLf & vbCrLf & _
           "Keep a copy on a different drive, so that losing this computer does not lose your library. " & _
           "The backup contains your private notes.", vbInformation, "Backup complete"
    CLNotifyManager
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbExclamation, "The backup was not completed"
End Sub

' ---------- restore ----------

' Restores into a new folder and connects to it. The restored library is given
' its own identity, so that if you later open the original folder Word tells
' you plainly that it is a different library rather than letting you work from
' a stale copy without noticing.
Public Function CLRestoreBackupTo(ByVal backup As String, ByVal parent As String, ByRef report As String) As String
    Dim marker As Object, folder As Variant, f As Object, copied As Long, destination As String
    Dim problems As String, e As Object, claimed As Long
    If Not CLExists(backup & "\backup-complete.xml") Then
        CLFail "That folder does not contain a finished Clause Library backup. Choose the folder named ""Clause Library backup ...""."
    End If
    Set marker = CLDom(CLRead(backup & "\backup-complete.xml"))
    If marker.documentElement.nodeName <> "backup" Or marker.documentElement.getAttribute("product") <> CL_PRODUCT Then
        CLFail "That folder does not contain a Clause Library backup."
    End If
    claimed = CLng(Val(marker.documentElement.getAttribute("files")))
    destination = parent & "\ClauseLibraryRestored-" & Format$(Now, "yyyy-mm-dd") & "-" & Left$(CLId(), 6)
    CLFolder destination: CLFolder destination & "\entries": CLFolder destination & "\history"
    For Each folder In Array("entries", "history")
        For Each f In CLFso().GetFolder(backup & "\" & folder).Files
            If CLBackupWanted(f.Name) Then
                On Error Resume Next
                Err.Clear
                CLFso().CopyFile f.path, destination & "\" & folder & "\" & f.Name, True
                If Err.number <> 0 Then problems = problems & "  " & f.Name & vbCrLf Else copied = copied + 1
                Err.Clear
                On Error GoTo 0
            End If
        Next
    Next
    Set e = CLDom(CLRead(backup & "\library.xml"))
    If e.documentElement.nodeName <> "library" Then CLFail "That backup's library marker is not valid, so nothing was restored."
    e.documentElement.setAttribute "id", CLId()
    e.documentElement.setAttribute "schema", CL_SCHEMA
    e.documentElement.setAttribute "restoredFrom", backup
    e.documentElement.setAttribute "restoredAt", CLStamp()
    ' Marker last, for the same reason as in backup.
    CLWrite destination & "\library.xml", e.XML
    report = copied & " of " & claimed & " files restored."
    If Len(problems) > 0 Then report = report & vbCrLf & vbCrLf & "These files could not be copied:" & vbCrLf & problems
    CLRestoreBackupTo = destination
End Function

Public Sub CLRestoreBackup()
    On Error GoTo Failed
    Dim picker As Object: Set picker = Application.FileDialog(4)
    picker.Title = "Choose the Clause Library backup folder to restore"
    If picker.Show <> -1 Then Exit Sub
    CLRequireLocalPath CStr(picker.SelectedItems(1)), "The folder you chose"
    If MsgBox("Restore this backup as a separate library, and use it from now on?" & vbCrLf & vbCrLf & _
              "Your current library folder is left exactly as it is. You can switch back at any time with Locate a library.", _
              vbOKCancel + vbQuestion, "Restore a backup") <> vbOK Then Exit Sub
    Dim report As String, restored As String
    restored = CLRestoreBackupTo(picker.SelectedItems(1), Environ$("LOCALAPPDATA"), report)
    CLRemember restored
    CLNotifyManager
    MsgBox "Restored and connected." & vbCrLf & vbCrLf & report & vbCrLf & vbCrLf & _
           "Your library is now:" & vbCrLf & restored & vbCrLf & vbCrLf & _
           "Your previous library folder was not touched.", vbInformation, "Restore complete"
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbExclamation, "Nothing was restored"
End Sub

Public Sub CLRibbonBackup(ByVal control As Object): CLBackUp: End Sub
Public Sub CLRibbonRestoreBackup(ByVal control As Object): CLRestoreBackup: End Sub
