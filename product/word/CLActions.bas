Attribute VB_Name = "CLActions"
' Clause Library - what the buttons do.
Option Explicit

Private manager As Object
Private ribbonUi As Object
Private draftWatcher As CLDraftWatcher
Private quickIds As Object          ' ribbon control id -> entry id

' ---------- readiness ----------

' Called before anything that touches the library. Offers the one-off format
' update if this library was made by version 1, taking a backup first.
Public Function CLEnsureReady() As Boolean
    Dim root As String
    root = CLRoot()
    CLSweepPending root & "\entries"
    If CLLibrarySchema(root) = CL_SCHEMA Then CLEnsureReady = True: Exit Function
    If MsgBox("Your library was created by an earlier version of Clause Library and needs a one-off update." & vbCrLf & vbCrLf & _
              "A complete backup will be made first, and nothing will be deleted. This usually takes a few seconds." & vbCrLf & vbCrLf & _
              "Update it now?", vbOKCancel + vbQuestion, "Clause Library") <> vbOK Then Exit Function
    Dim backup As String, report As String
    backup = CLBackupTo(Environ$("LOCALAPPDATA"), True)
    CLMigrate root, report
    MsgBox "Your library is ready." & vbCrLf & vbCrLf & report & vbCrLf & vbCrLf & _
           "A backup of the previous format was saved to:" & vbCrLf & backup, vbInformation, "Clause Library"
    CLEnsureReady = True
End Function

' ---------- the library window ----------

Public Sub ClauseLibrary()
    On Error GoTo Failed
    If Not CLEnsureReady() Then Exit Sub
    If manager Is Nothing Then Set manager = VBA.UserForms.Add("frmLibrary")
    manager.Show vbModeless
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbInformation, "Clause Library"
End Sub

Public Sub CLReleaseManager()
    Set manager = Nothing
End Sub

Public Function CLCloseManager() As Boolean
    If Not manager Is Nothing Then Unload manager
    CLCloseManager = (manager Is Nothing)
End Function

Public Sub CLNotifyManager()
    On Error Resume Next
    If Not manager Is Nothing Then manager.RefreshEntries
    CLRefreshRibbon
End Sub

' A visible, lasting confirmation. The status bar alone is too easy to miss,
' and a stale status bar message reads exactly like a fresh one.
Public Sub CLSay(ByVal message As String)
    Application.StatusBar = message
    If Not manager Is Nothing Then
        On Error Resume Next
        manager.SetStatus message
        On Error GoTo 0
    End If
End Sub

' ---------- capture ----------

Public Sub CLCaptureSelection()
    On Error GoTo Failed
    If Documents.count = 0 Then CLFail "Open a document and select the wording you want to keep."
    If Not CLEnsureReady() Then Exit Sub
    Dim source As Range: Set source = Selection.Range.Duplicate
    If source.Start = source.End Then CLFail "Select the wording you want to keep, then choose Capture selection."
    If source.StoryType = wdCommentsStory Then CLCaptureComment: Exit Sub

    Dim e As Object, plain As String, xml As String, notes As String, title As String
    Dim c As Comment, commentCount As Long, fingerprint As String, existing As Object

    If source.Revisions.count > 0 Then
        If MsgBox("This selection contains tracked changes." & vbCrLf & vbCrLf & _
                  "Save the final wording, with those changes treated as accepted in the library copy only?" & vbCrLf & vbCrLf & _
                  "Your document and its tracked changes are not altered.", _
                  vbOKCancel + vbQuestion, "Capture final wording") <> vbOK Then Exit Sub
    End If

    xml = CLCapturePayload(source, True, plain)

    fingerprint = CLFingerprint(plain)
    Set existing = CLFindFingerprint(CLList(), fingerprint)
    If Not existing Is Nothing Then
        If MsgBox("You already have this wording in your library:" & vbCrLf & vbCrLf & _
                  "    " & existing("title") & vbCrLf & vbCrLf & _
                  "Save it again as a separate item?", vbYesNo + vbQuestion + vbDefaultButton2, "Already in your library") <> vbYes Then
            CLShowEntry existing("id")
            Exit Sub
        End If
    End If

    commentCount = source.Comments.count
    If commentCount > 0 Then notes = CLAskAboutComments(source, commentCount)

    title = CLSuggestTitle(source, plain)
    If CLSetting("PromptForTitle", "1") = "1" Then
        Dim answer As String
        answer = InputBox("Name this clause so you can find it later. Press Enter to accept the suggestion." & vbCrLf & vbCrLf & _
                          "(You can rename it at any time in Library.)", "Save to Clause Library", title)
        If StrPtr(answer) <> 0 Then
            If Len(Trim$(answer)) > 0 Then title = Trim$(answer)
        End If
    End If

    Set e = CLNew("Clause")
    CLSet e, "title", Left$(title, 120)
    CLSet e, "plain", plain
    CLSet e, "preview", CLCapturePreview
    CLSet e, "notes", notes
    CLSet e, "source", source.Document.Name
    CLSet e, "fingerprint", fingerprint
    CLSave e, 0, xml, True

    CLSay "Saved to Clause Library: " & Left$(title, 60)
    CLNotifyManager
    Dim extra As String: extra = CLCaptureNotes
    If Len(extra) > 0 Then
        MsgBox "Saved to your library as:" & vbCrLf & vbCrLf & "    " & title & vbCrLf & vbCrLf & extra, _
               vbInformation, "Saved - please check"
    End If
    Exit Sub
Failed:
    MsgBox "Nothing was saved. " & CLExplain(Err.number, Err.Description), vbExclamation, "Capture selection"
End Sub

' Comments on wording in a draft you received belong to whoever wrote them.
' They are never taken into the library without being asked for.
Private Function CLAskAboutComments(ByVal source As Range, ByVal commentCount As Long) As String
    Dim c As Comment, notes As String, preview As String
    For Each c In source.Comments
        If Len(notes) > 0 Then notes = notes & vbCrLf & vbCrLf
        notes = notes & Trim$(CLSafeText(c.Range.Text))
    Next
    preview = Left$(notes, 300): If Len(notes) > 300 Then preview = preview & "..."
    If MsgBox("This wording has " & commentCount & " comment" & IIf(commentCount = 1, "", "s") & " attached to it:" & vbCrLf & vbCrLf & _
              preview & vbCrLf & vbCrLf & _
              "Keep " & IIf(commentCount = 1, "it", "them") & " as your private notes on this item?" & vbCrLf & vbCrLf & _
              "Private notes are never inserted into a contract, but they are stored in your library and included in backups. " & _
              "If these comments are the other side's, or privileged, choose No.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Comments on this wording") = vbYes Then
        CLAskAboutComments = notes
    End If
End Function

' A name you would actually recognise in a list: the heading this wording sits
' under, where there is one, otherwise its opening words cut at a word break.
Public Function CLSuggestTitle(ByVal source As Range, ByVal plain As String) As String
    Dim heading As String, opening As String, p As Paragraph, guard As Long
    On Error Resume Next
    Set p = source.Paragraphs(1)
    For guard = 1 To 30
        If p Is Nothing Then Exit For
        If p.OutlineLevel <> wdOutlineLevelBodyText Then
            heading = Trim$(CLSafeText(p.Range.Text))
            Exit For
        End If
        If p.Range.Start <= 0 Then Exit For
        Set p = p.Previous
    Next
    On Error GoTo 0
    opening = CLFirstWords(plain, 60)
    If Len(heading) > 0 Then
        heading = CLFirstWords(heading, 48)
        If InStr(1, opening, heading, vbTextCompare) = 1 Then
            CLSuggestTitle = opening
        Else
            CLSuggestTitle = heading & " - " & CLFirstWords(plain, 40)
        End If
    Else
        CLSuggestTitle = opening
    End If
    CLSuggestTitle = Trim$(CLSuggestTitle)
    If Len(CLSuggestTitle) = 0 Then CLSuggestTitle = "Untitled clause"
End Function

Private Function CLFirstWords(ByVal value As String, ByVal width As Long) As String
    Dim t As String, cut As Long
    t = Trim$(Replace$(Replace$(CLSafeText(value), vbLf, " "), vbTab, " "))
    Do While InStr(t, "  ") > 0: t = Replace$(t, "  ", " "): Loop
    ' Drop a leading clause number so the name is the wording, not "14.3".
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[\(\[]?[0-9]+(\.[0-9]+)*[\)\]\.]?\s+"
    t = re.Replace(t, "")
    If Len(t) <= width Then CLFirstWords = t: Exit Function
    cut = InStrRev(Left$(t, width + 1), " ")
    If cut < width \ 2 Then cut = width
    CLFirstWords = RTrim$(Left$(t, cut))
End Function

Public Sub CLCaptureComment()
    On Error GoTo Failed
    If Documents.count = 0 Then CLFail "Open the document containing the comment first."
    If Not CLEnsureReady() Then Exit Sub
    Dim plain As String, e As Object, title As String
    If Selection.Range.StoryType = wdCommentsStory Then
        plain = CLPlainWording(Selection.Range.Text)
    ElseIf Selection.Comments.count = 1 Then
        plain = CLPlainWording(Selection.Comments(1).Range.Text)
    Else
        CLFail "Select the text of one comment, or select the wording in the document that has exactly one comment attached."
    End If
    If Len(Trim$(plain)) = 0 Then CLFail "There is no comment wording in that selection."
    title = CLFirstWords(plain, 60)
    If CLSetting("PromptForTitle", "1") = "1" Then
        Dim answer As String
        answer = InputBox("Name this comment. Press Enter to accept the suggestion.", "Save comment to Clause Library", title)
        If StrPtr(answer) <> 0 Then
            If Len(Trim$(answer)) > 0 Then title = Trim$(answer)
        End If
    End If
    Set e = CLNew("Comment")
    CLSet e, "plain", plain: CLSet e, "comment", plain
    CLSet e, "title", Left$(title, 120)
    CLSet e, "source", ActiveDocument.Name
    CLSet e, "fingerprint", CLFingerprint(plain)
    CLSave e, 0
    CLSay "Comment saved to Clause Library: " & Left$(title, 60)
    CLNotifyManager
    Exit Sub
Failed:
    MsgBox "Nothing was saved. " & CLExplain(Err.number, Err.Description), vbExclamation, "Capture comment"
End Sub

' ---------- reuse ----------

Public Sub CLInsertEntry(ByVal id As String, Optional ByVal plainOnly As Boolean = False, _
                         Optional ByVal reviewedRevision As Long = -1)
    On Error GoTo Failed
    Dim e As Object, status As String, xml As String, target As Range, inserted As Range
    Dim blocked As String, mismatch As String, plain As String
    Set e = CLLoad(id, status)
    ' The wording a lawyer approved with their eyes must be the wording that
    ' lands in the contract. If the record moved on since it was displayed,
    ' stop and make them look again rather than quietly inserting something else.
    If reviewedRevision >= 0 And CLRevision(e) <> reviewedRevision Then
        CLReselect id
        CLFail "This item changed since the wording you were shown was loaded." & vbCrLf & vbCrLf & _
               "It has been reloaded - version " & reviewedRevision & " became version " & CLRevision(e) & "." & vbCrLf & vbCrLf & _
               "Nothing was inserted. Read the current wording, then insert it."
    End If
    If CLGet(e, "state") <> "Active" Then CLFail "This item is in Archive or Trash. Restore it in Library before using it."
    If CLGet(e, "kind") = "Comment" Then CLInsertComment id, reviewedRevision: Exit Sub
    If Documents.count = 0 Then CLFail "Open the document you want to insert into first."
    Set target = Selection.Range.Duplicate
    plain = CLGet(e, "plain")

    If target.Start <> target.End Then
        If MsgBox("Replace the selected wording in " & target.Document.Name & "?", _
                  vbOKCancel + vbQuestion, "Insert clause") <> vbOK Then Exit Sub
    End If

    xml = CLReadPayload(e)
    If Len(xml) = 0 Then plainOnly = True
    If Not plainOnly Then
        blocked = CLInsertBlockedReason(target, False)
        If Len(blocked) > 0 Then
            If InStr(blocked, "Insert plain text") = 0 Then CLFail blocked
            If MsgBox(blocked & vbCrLf & vbCrLf & "Insert the wording as plain text here instead?", _
                      vbOKCancel + vbQuestion, "Insert clause") <> vbOK Then Exit Sub
            plainOnly = True
        End If
    End If
    ' Plain insertion used to skip every destination check the formatted path
    ' makes. It is simpler, not less consequential, so it goes through the same
    ' preflight - just with table cells allowed, where plain text is fine.
    If plainOnly Then
        blocked = CLInsertBlockedReason(target, True)
        If Len(blocked) > 0 Then CLFail blocked
    End If
    If Not plainOnly Then
        mismatch = CLPayloadMismatch(xml, plain)
        If Len(mismatch) > 0 Then
            If MsgBox("The saved formatting for this item no longer produces exactly the wording recorded with it (" & mismatch & ")." & vbCrLf & vbCrLf & _
                      "This usually means Word has changed how it stores content since the item was saved." & vbCrLf & vbCrLf & _
                      "Insert the saved formatted version anyway?" & vbCrLf & _
                      "Choose No to insert the recorded wording as plain text instead.", _
                      vbYesNo + vbExclamation, "Check this wording") <> vbYes Then plainOnly = True
        End If
    End If

    If plainOnly Then
        Dim startAt As Long: startAt = target.Start
        Dim originalTracking As Boolean: originalTracking = target.Document.TrackRevisions
        Application.UndoRecord.StartCustomRecord "Insert Clause Library wording"
        target.Text = Replace$(plain, vbLf, vbCr)
        Application.UndoRecord.EndCustomRecord
        target.Document.TrackRevisions = originalTracking
        Set inserted = target.Document.Range(startAt, target.End)
    Else
        CLInsertPayload target, xml, target.Document.TrackRevisions, inserted
    End If

    CLTouchUsage id
    CLNotifyManager
    If CLSelectFirstPlaceholder(inserted) Then
        CLSay "Inserted. The first placeholder is selected - complete it before circulating."
    Else
        If Not inserted Is Nothing Then inserted.Select
        CLSay "Inserted from Clause Library. Check defined terms and the fit with this contract."
    End If
    Exit Sub
Failed:
    Dim message As String: message = CLExplain(Err.number, Err.Description)
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    MsgBox message, vbExclamation, "Nothing was inserted"
End Sub

Public Sub CLInsertComment(ByVal id As String, Optional ByVal reviewedRevision As Long = -1)
    On Error GoTo Failed
    Dim e As Object: Set e = CLLoad(id)
    If reviewedRevision >= 0 And CLRevision(e) <> reviewedRevision Then
        CLReselect id
        CLFail "This comment changed since it was shown to you. It has been reloaded; nothing was added. Read it, then add it."
    End If
    If CLGet(e, "state") <> "Active" Then CLFail "This comment is in Archive or Trash. Restore it in Library before using it."
    Dim text As String: text = CLGet(e, "comment")
    If Len(text) = 0 Then CLFail "This item has no reusable comment text. Private notes are never inserted."
    If Documents.count = 0 Then CLFail "Open a document first."
    If Selection.Range.Start = Selection.Range.End Then CLFail "Select the wording this comment should be attached to."
    If ActiveDocument.ReadOnly Or ActiveDocument.ProtectionType <> wdNoProtection Then
        CLFail "This document is read-only or protected, so no comment was added."
    End If
    If MsgBox("Add this comment to " & ActiveDocument.Name & "?" & vbCrLf & vbCrLf & _
              Left$(text, 500) & IIf(Len(text) > 500, "...", "") & vbCrLf & vbCrLf & _
              "Anyone you send this document to will be able to read it.", _
              vbOKCancel + vbQuestion, "Add a comment others will see") <> vbOK Then Exit Sub
    ActiveDocument.Comments.Add Range:=Selection.Range.Duplicate, Text:=text
    CLTouchUsage id
    CLNotifyManager
    CLSay "Comment added to " & ActiveDocument.Name & "."
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbExclamation, "No comment was added"
End Sub

' ---------- editing wording ----------

Public Sub CLEditWording(ByVal id As String)
    On Error GoTo Failed
    Dim e As Object, d As Document, xml As String, title As String
    Set e = CLLoad(id)
    title = CLGet(e, "title")
    xml = CLReadPayload(e)
    Set d = Documents.Add
    d.TrackRevisions = False
    If Len(xml) > 0 Then
        CLValidatePayload xml
        d.Content.InsertXML xml
    Else
        d.Content.Text = Replace$(CLGet(e, "plain"), vbLf, vbCr)
    End If
    CLWatchDrafts
    d.Variables.Add "ClauseLibraryEntry", id
    d.Variables.Add "ClauseLibraryRevision", CStr(CLRevision(e))
    d.Variables.Add "ClauseLibraryRoot", CLRoot()
    d.BuiltInDocumentProperties("Title") = "Clause Library draft: " & title
    On Error Resume Next
    d.Windows(1).Caption = "CLAUSE LIBRARY DRAFT - " & Left$(title, 60) & " (not a contract)"
    On Error GoTo Failed
    d.Saved = True
    CLSay "Library draft open. Edit it, then choose Save wording on the Clause Library tab. Closing it does not change your library."
    MsgBox "This is a working draft of your library wording, not a contract." & vbCrLf & vbCrLf & _
           "Edit it as you would any document, then choose Save wording on the Clause Library tab." & vbCrLf & vbCrLf & _
           "If you close it without saving, your library is unchanged.", vbInformation, "Editing: " & Left$(title, 60)
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbExclamation, "The draft could not be opened"
End Sub

' Word only tells us about saves and closes while something is listening.
Public Sub CLWatchDrafts()
    On Error Resume Next
    If draftWatcher Is Nothing Then
        Set draftWatcher = New CLDraftWatcher
        draftWatcher.Watch
    End If
    On Error GoTo 0
End Sub

Public Sub CLSaveWording()
    If Documents.count = 0 Then
        MsgBox "Open a Clause Library wording draft first.", vbInformation, "Save wording"
        Exit Sub
    End If
    CLSaveWordingFor ActiveDocument
End Sub

' Returns True only if the library really was updated.
Public Function CLSaveWordingFor(ByVal d As Document) As Boolean
    On Error GoTo Failed
    Dim id As String, rootOfDraft As String, expected As Long
    id = CLDocVariable(d, "ClauseLibraryEntry")
    If Len(id) = 0 Then
        CLFail "This document is not a Clause Library wording draft." & vbCrLf & vbCrLf & _
               "To change saved wording, open Library, select the item and choose Edit wording in Word."
    End If
    rootOfDraft = CLDocVariable(d, "ClauseLibraryRoot")
    If rootOfDraft <> CLRoot() Then CLFail "This draft belongs to a different library folder than the one now connected. It was not saved."
    expected = CLng(Val(CLDocVariable(d, "ClauseLibraryRevision")))
    Dim e As Object, plain As String, xml As String
    Set e = CLLoad(id)
    If CLRevision(e) <> expected Then
        CLFail "This item has changed since you opened this draft, so it was not overwritten." & vbCrLf & vbCrLf & _
               "Save this document somewhere of your own if you want to keep your edits, then open the item again from Library."
    End If
    If CLGet(e, "kind") = "Comment" Then
        plain = CLPlainWording(d.Content.Text)
        CLSet e, "comment", plain
        CLSet e, "plain", plain
        CLSet e, "fingerprint", CLFingerprint(plain)
        CLSave e, expected
        If MsgBox("Saved." & vbCrLf & vbCrLf & "Do the topic, tags and private notes on this item still describe it correctly?" & vbCrLf & vbCrLf & _
                  "Choose No to mark it as needing a look.", vbYesNo + vbQuestion, "Details") = vbNo Then CLMarkUnorganised id
    Else
        If d.Revisions.count > 0 Then
            If MsgBox("This draft contains tracked changes." & vbCrLf & vbCrLf & _
                      "Save the final wording, with those changes treated as accepted?", _
                      vbOKCancel + vbQuestion, "Save wording") <> vbOK Then Exit Sub
        End If
        xml = CLCapturePayload(d.Content, True, plain)
        CLSet e, "plain", plain
        CLSet e, "preview", CLCapturePreview
        CLSet e, "fingerprint", CLFingerprint(plain)
        CLSave e, expected, xml, True
        If MsgBox("Saved." & vbCrLf & vbCrLf & "Do the topic, tags, Use when and private notes on this item still describe it correctly?" & vbCrLf & vbCrLf & _
                  "Choose No to mark it as needing a look. Nothing stops you using it either way.", _
                  vbYesNo + vbQuestion, "Details") = vbNo Then CLMarkUnorganised id
    End If
    d.Variables("ClauseLibraryRevision").Value = CStr(CLRevision(e))
    d.Saved = True
    CLSaveWordingFor = True
    CLSay "Wording saved. The previous version is kept under Previous versions."
    CLNotifyManager
    MsgBox "Saved to your library as version " & CLRevision(e) & "." & vbCrLf & vbCrLf & _
           "The version you replaced is still available under Previous versions.", _
           vbInformation, "Save wording"
    Exit Function
Failed:
    MsgBox "Your library was not changed." & vbCrLf & vbCrLf & CLExplain(Err.number, Err.Description), _
           vbExclamation, "Save wording"
End Function

' Puts an item back on the Not organised list, because its wording changed and
' the description of it may no longer be true. Never blocks reuse.
Public Sub CLMarkUnorganised(ByVal id As String)
    On Error Resume Next
    Dim e As Object: Set e = CLLoad(id)
    If e Is Nothing Then Exit Sub
    If CLGet(e, "organised") <> "1" Then Exit Sub
    CLSet e, "organised", "0"
    CLSave e, CLRevision(e)
    CLNotifyManager
End Sub

Private Function CLDocVariable(ByVal d As Document, ByVal name As String) As String
    On Error Resume Next
    CLDocVariable = d.Variables(name).Value
    On Error GoTo 0
End Function

' ---------- library location ----------

Public Sub CLLocateLibrary()
    On Error GoTo Failed
    Dim picker As Object: Set picker = Application.FileDialog(4)
    picker.Title = "Choose your Clause Library folder"
    If picker.Show <> -1 Then Exit Sub
    CLRemember picker.SelectedItems(1)
    CLNotifyManager
    MsgBox "Connected to:" & vbCrLf & vbCrLf & picker.SelectedItems(1) & vbCrLf & vbCrLf & _
           "This is the library Word will use from now on, including after a restart.", vbInformation, "Clause Library"
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbExclamation, "That library was not connected"
End Sub

' Refresh and re-select, but only in a window that is already open.
Public Sub CLReselect(ByVal id As String)
    On Error Resume Next
    If manager Is Nothing Then Exit Sub
    manager.RefreshEntries
    manager.SelectEntry id
End Sub

Public Sub CLShowEntry(ByVal id As String)
    On Error Resume Next
    ClauseLibrary
    If Not manager Is Nothing Then manager.SelectEntry id
End Sub

' ---------- help ----------

Public Sub CLHelp()
    MsgBox "CAPTURE" & vbCrLf & _
      "Select wording in any document and choose Capture selection. Your document is never changed. Give it a name when asked - that is what you will search for later." & vbCrLf & vbCrLf & _
      "FIND AND REUSE" & vbCrLf & _
      "Insert recent gives you your last and favourite clauses in one click. Library searches everything, including your private notes. Put the cursor where the wording should go, then Insert." & vbCrLf & vbCrLf & _
      "ORGANISE WHEN IT SUITS YOU" & vbCrLf & _
      "Topic, tags, family and notes are all optional and never stop you reusing something. Not organised is a reminder, nothing more." & vbCrLf & vbCrLf & _
      "EDIT" & vbCrLf & _
      "Edit wording in Word opens a draft. Change it, then choose Save wording. Every previous version is kept." & vbCrLf & vbCrLf & _
      "KEEP IT SAFE" & vbCrLf & _
      "Back up library writes a checked copy wherever you choose - use another drive. Export writes your clauses out as ordinary Word files you can read without this tool.", _
      vbInformation, "Clause Library - quick guide"
End Sub

Public Sub CLAbout()
    Dim root As String, count As String, lastBackup As String
    On Error Resume Next
    root = CLRoot()
    count = CLList().count & " items"
    lastBackup = CLSetting("LastBackup", "never")
    On Error GoTo 0
    MsgBox "Clause Library " & CL_VERSION & vbCrLf & _
           "Library format " & CL_SCHEMA & "  |  Word " & Application.Version & vbCrLf & vbCrLf & _
           "Connected library:" & vbCrLf & IIf(Len(root) = 0, "(none yet)", root) & vbCrLf & vbCrLf & _
           "Contents: " & count & vbCrLf & _
           "Last backup: " & lastBackup & vbCrLf & vbCrLf & _
           "Everything is stored on this computer. Nothing is sent anywhere.", _
           vbInformation, "About Clause Library"
End Sub

' ---------- ribbon ----------

Public Sub CLRibbonLoad(ByVal ribbon As Object)
    Set ribbonUi = ribbon
End Sub

Public Sub CLRefreshRibbon()
    On Error Resume Next
    If Not ribbonUi Is Nothing Then ribbonUi.Invalidate
End Sub

' Builds the drop-down of wording you actually reach for. This is the one-click
' path that makes the library part of drafting rather than a place to visit.
Public Function CLQuickMenu(ByVal control As Object) As String
    On Error GoTo NoLibrary
    Dim rows As Collection, r As Variant, xml As String, n As Long, sortBy As String, label As String
    Dim wanted As String: wanted = control.id
    Set rows = CLList()
    If quickIds Is Nothing Then Set quickIds = CreateObject("Scripting.Dictionary")
    quickIds.RemoveAll
    sortBy = IIf(wanted = "CLFavouritesMenu", "Title", "Recently used")
    Set rows = CLSearch(rows, "", IIf(wanted = "CLFavouritesMenu", "Favourites", "All"), sortBy)
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"
    For Each r In rows
        If wanted = "CLFavouritesMenu" Or Len(r("lastUsed")) > 0 Then
            n = n + 1
            If n > 15 Then Exit For
            quickIds("CLQuick" & n) = r("id")
            label = CLHtml(Left$(r("title"), 70))
            xml = xml & "<button id=""CLQuick" & n & """ label=""" & label & """ onAction=""CLRibbonQuickInsert"" screentip=""" & _
                  CLHtml(r("kind") & IIf(Len(r("topic")) > 0, " - " & r("topic"), "")) & """/>"
        End If
    Next
    If n = 0 Then
        xml = xml & "<button id=""CLQuickNone"" label=""" & _
              IIf(wanted = "CLFavouritesMenu", "No favourites yet", "Nothing used yet") & """ enabled=""false""/>"
    End If
    xml = xml & "<menuSeparator id=""CLQuickSep""/>" & _
          "<button id=""CLQuickOpen"" label=""Open Library..."" imageMso=""FileOpen"" onAction=""CLRibbonOpen""/></menu>"
    CLQuickMenu = xml
    Exit Function
NoLibrary:
    CLQuickMenu = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
                  "<button id=""CLQuickSetup"" label=""Set up Clause Library first"" enabled=""false""/></menu>"
End Function

Public Sub CLRibbonQuickInsert(ByVal control As Object)
    On Error GoTo Failed
    If quickIds Is Nothing Then Exit Sub
    If Not quickIds.Exists(control.id) Then Exit Sub
    CLInsertEntry CStr(quickIds(control.id))
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbExclamation, "Nothing was inserted"
End Sub

Public Sub CLRibbonOpen(ByVal control As Object): ClauseLibrary: End Sub
Public Sub CLRibbonCapture(ByVal control As Object): CLCaptureSelection: End Sub
Public Sub CLRibbonCaptureComment(ByVal control As Object): CLCaptureComment: End Sub
Public Sub CLRibbonSave(ByVal control As Object): CLSaveWording: End Sub
Public Sub CLRibbonHelp(ByVal control As Object): CLHelp: End Sub
Public Sub CLRibbonAbout(ByVal control As Object): CLAbout: End Sub
Public Sub CLRibbonLocate(ByVal control As Object): CLLocateLibrary: End Sub
