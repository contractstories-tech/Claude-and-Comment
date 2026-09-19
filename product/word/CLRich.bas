Attribute VB_Name = "CLRich"
' Clause Library - the Word formatting engine.
'
' Two rules govern everything here:
'   the document you captured from is never modified, and
'   content that cannot be reused safely is refused out loud, not flattened quietly.
'
' Capture resolves tracked changes, comments and fields in a throwaway copy, so
' the clause you keep is the wording as finally agreed, and your draft is untouched.
Option Explicit

Private Const WNS As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
Private Const PKGNS As String = "http://schemas.microsoft.com/office/2006/xmlPackage"
Private Const RELNS As String = "http://schemas.openxmlformats.org/package/2006/relationships"

' Filled in by CLCapturePayload so the caller can tell the person what happened
' to their selection rather than leaving them to discover it later.
Public CLCaptureNotes As String
Public CLFieldsFrozen As Long
' Wording with its list numbers put back in, for the preview panel. Range.Text
' leaves automatic numbering out, so a preview built from it misleads.
Public CLCapturePreview As String

Public Function CLPlainWording(ByVal value As String) As String
    CLPlainWording = CLTrimWording(CLSafeText(value))
End Function

Private Function XmlDocument(ByVal xml As String) As Object
    Dim dom As Object: Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False: dom.validateOnParse = False: dom.resolveExternals = False
    dom.setProperty "ProhibitDTD", True
    dom.setProperty "SelectionNamespaces", "xmlns:w='" & WNS & "' xmlns:pkg='" & PKGNS & "' xmlns:r='" & RELNS & "'"
    If Not dom.LoadXML(xml) Then CLFail "The saved Word formatting for this item is not valid and was not used."
    Set XmlDocument = dom
End Function

' Refuses anything that would carry revision history, other people's comments,
' live fields, external links or active content into or out of the library.
Public Sub CLValidatePayload(ByVal xml As String)
    If Len(xml) = 0 Then CLFail "There is no saved Word formatting for this item."
    Dim dom As Object: Set dom = XmlDocument(xml)
    If dom.selectNodes("/pkg:package/pkg:part[@pkg:name='/word/document.xml']").Length <> 1 Then
        CLFail "This saved formatting is not a single piece of Word content and was not used."
    End If
    If dom.selectNodes("//w:ins | //w:del | //w:moveFrom | //w:moveTo | //w:moveFromRangeStart | //w:moveToRangeStart | //w:cellIns | //w:cellDel | //w:cellMerge | //w:pPrChange | //w:rPrChange | //w:tblPrChange | //w:trPrChange | //w:tcPrChange | //w:sectPrChange").Length > 0 Then
        CLFail "This saved formatting still contains tracked changes, so it was not used. Capture the wording again."
    End If
    If dom.selectNodes("//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference | //w:comment").Length > 0 Then
        CLFail "This saved formatting still contains comments, so it was not used. Capture the wording again."
    End If
    If dom.selectNodes("//w:instrText | //w:fldChar | //w:fldSimple | //w:altChunk | //w:object | //w:dataBinding | //w:subDoc").Length > 0 Then
        CLFail "This saved formatting contains fields, linked objects or embedded content that cannot be reused safely. It was not used."
    End If
    If dom.selectNodes("//r:Relationship[translate(@TargetMode,'EXTERNAL','external')='external']").Length > 0 Then
        CLFail "This saved formatting depends on a file or address outside your library, so it was not used."
    End If
    If dom.selectNodes("/pkg:package/pkg:part[contains(@pkg:name,'vbaProject') or contains(@pkg:name,'embeddings/') or contains(@pkg:name,'customXml/') or contains(@pkg:name,'comments')]").Length > 0 Then
        CLFail "This saved formatting contains macros or private data parts and was not used."
    End If
End Sub

' ---------- capture ----------

Public Function CLCapturePayload(ByVal source As Range, ByVal acceptFinal As Boolean, ByRef plain As String) As String
    Dim temp As Document, raw As String, c As Long, message As String, table As Table
    Dim notes As String, removedHidden As Long
    On Error GoTo Failed
    CLCaptureNotes = "": CLFieldsFrozen = 0
    If source.Start = source.End Then CLFail "Select the wording you want to keep first."
    If source.StoryType <> wdMainTextStory Then
        CLFail "Rich capture works on wording in the body of the document. Select the wording there, or use Capture comment for a comment."
    End If
    If source.Revisions.count > 0 And Not acceptFinal Then
        CLFail "This selection contains tracked changes. Choose to keep the final wording, or cancel. Your document's changes are never altered."
    End If
    For Each table In source.Tables
        If source.Start > table.Range.Start Or source.End < table.Range.End Then
            CLFail "Part of a table is selected. Select the whole table, or select wording outside it. Nothing is flattened silently."
        End If
    Next
    raw = source.WordOpenXML
    Set temp = Documents.Add(Visible:=False)
    temp.TrackRevisions = False
    temp.Content.InsertXML raw
    If temp.Revisions.count > 0 Then temp.Revisions.AcceptAll
    For c = temp.Comments.count To 1 Step -1: temp.Comments(c).Delete: Next
    For c = temp.Hyperlinks.count To 1 Step -1: temp.Hyperlinks(c).Delete: Next
    ' Cross-references and other fields become the text they currently show.
    ' That is the only way to store them, and the person is told, because a
    ' frozen "clause 14.3" is wrong in the next contract.
    CLFieldsFrozen = temp.Fields.count
    For c = temp.Fields.count To 1 Step -1: temp.Fields(c).Unlink: Next
    For c = 1 To temp.InlineShapes.count
        If temp.InlineShapes(c).Type <> wdInlineShapePicture Then
            CLFail "This selection contains an embedded or linked object, which cannot be reused as wording. Remove it from the selection and capture again."
        End If
    Next
    If temp.Shapes.count > 0 Then
        CLFail "This selection contains a floating picture or text box, which cannot be reused as wording. Remove it from the selection and capture again."
    End If
    removedHidden = CLRemoveHidden(temp)
    plain = CLPlainWording(temp.Content.Text)
    If Len(Trim$(plain)) = 0 Then CLFail "There is no reusable wording in this selection once tracked changes and comments are resolved."
    CLCapturePreview = CLNumberedText(temp.Content)
    raw = temp.Content.WordOpenXML
    ' Page setup belongs to the destination contract, never to a clause.
    Dim dom As Object, n As Object: Set dom = XmlDocument(raw)
    For Each n In dom.selectNodes("//w:sectPr"): n.parentNode.removeChild n: Next
    raw = dom.XML
    CLValidatePayload raw
    If CLFieldsFrozen > 0 Then
        notes = CLFieldsFrozen & " field" & IIf(CLFieldsFrozen = 1, "", "s") & _
                " (such as cross-references or dates) were saved as the text they showed. Check them after inserting."
    End If
    If removedHidden > 0 Then
        If Len(notes) > 0 Then notes = notes & vbCrLf
        notes = notes & "Hidden text was left out of the saved wording."
    End If
    CLCaptureNotes = notes
    CLCapturePayload = raw
    temp.Close wdDoNotSaveChanges
    Exit Function
Failed:
    message = CLExplain(Err.number, Err.Description)
    On Error Resume Next
    If Not temp Is Nothing Then temp.Close wdDoNotSaveChanges
    On Error GoTo 0
    CLFail message
End Function

' Hidden text is dropped, but the loop is bounded: a find that cannot be
' deleted would otherwise restart for ever and Word would appear to freeze.
Private Function CLRemoveHidden(ByVal temp As Document) As Long
    Dim hidden As Range, guard As Long, before As Long, removed As Long
    Set hidden = temp.Content.Duplicate
    Do While guard < 500
        guard = guard + 1
        Set hidden = temp.Content.Duplicate
        With hidden.Find
            .ClearFormatting: .Text = "": .Font.hidden = True
            .Forward = True: .Wrap = wdFindStop: .Format = True
        End With
        If Not hidden.Find.Execute Then Exit Do
        before = temp.Content.End
        hidden.Delete
        If temp.Content.End >= before Then Exit Do      ' nothing was removed; stop rather than spin
        removed = removed + 1
    Loop
    CLRemoveHidden = removed
End Function

' ---------- insertion ----------

' Answers whether rich insertion is possible at this position, and why not.
Public Function CLInsertBlockedReason(ByVal target As Range) As String
    Dim doc As Document, tbl As Table
    Set doc = target.Document
    If doc.ReadOnly Then CLInsertBlockedReason = "This document is read-only, so nothing was inserted.": Exit Function
    If doc.ProtectionType <> wdNoProtection Then CLInsertBlockedReason = "This document is protected, so nothing was inserted. Remove the protection and try again.": Exit Function
    If target.StoryType <> wdMainTextStory Then CLInsertBlockedReason = "Place the cursor in the body of the document.": Exit Function
    If target.Revisions.count > 0 Then
        CLInsertBlockedReason = "The wording you are replacing contains tracked changes. Accept or reject them first, so it is clear what is being replaced."
        Exit Function
    End If
    If target.Information(wdWithInTable) And target.Tables.count = 0 Then
        CLInsertBlockedReason = "Formatted wording cannot be placed inside a table cell. Use Insert plain text here, or put the cursor outside the table."
        Exit Function
    End If
    For Each tbl In doc.Tables
        If target.Start >= tbl.Range.Start And target.Start < tbl.Range.End Then
            If target.Start <> tbl.Range.Start Or target.End < tbl.Range.End Then
                CLInsertBlockedReason = "Formatted wording cannot be placed inside part of a table. Use Insert plain text, select the whole table, or put the cursor outside it."
                Exit Function
            End If
        ElseIf target.End > tbl.Range.Start And target.End < tbl.Range.End Then
            CLInsertBlockedReason = "Your selection ends inside a table. Select the whole table, or select wording outside it."
            Exit Function
        End If
    Next
End Function

' Checks the saved formatting really does produce the wording this entry
' records. Returns a description of any difference rather than refusing: Word
' changes how it normalises content between versions, and a library should
' still open years later.
Public Function CLPayloadMismatch(ByVal xml As String, ByVal expectedText As String) As String
    Dim probe As Document, produced As String
    On Error GoTo Failed
    Set probe = Documents.Add(Visible:=False)
    probe.TrackRevisions = False
    probe.Content.InsertXML xml
    produced = CLPlainWording(probe.Content.Text)
    probe.Close wdDoNotSaveChanges: Set probe = Nothing
    If produced = expectedText Then Exit Function
    If CLNormalise(produced) = CLNormalise(expectedText) Then
        CLPayloadMismatch = "spacing or punctuation only"
    Else
        CLPayloadMismatch = "the wording itself differs"
    End If
    Exit Function
Failed:
    On Error Resume Next
    If Not probe Is Nothing Then probe.Close wdDoNotSaveChanges
    On Error GoTo 0
    CLPayloadMismatch = "the saved formatting could not be opened"
End Function

' Inserts the payload at target, inside one Undo step. On any failure the
' document is put back exactly as it was.
Public Sub CLInsertPayload(ByVal target As Range, ByVal xml As String, ByVal trackInsertion As Boolean, ByRef inserted As Range)
    Dim doc As Document, originalTracking As Boolean, undoStarted As Boolean
    Dim attempted As Boolean, startAt As Long, message As String
    On Error GoTo Failed
    Set doc = target.Document
    originalTracking = doc.TrackRevisions
    CLValidatePayload xml
    message = CLInsertBlockedReason(target)
    If Len(message) > 0 Then CLFail message
    startAt = target.Start
    doc.Activate
    doc.TrackRevisions = trackInsertion
    Application.UndoRecord.StartCustomRecord "Insert Clause Library wording": undoStarted = True
    attempted = True
    target.InsertXML xml
    Application.UndoRecord.EndCustomRecord: undoStarted = False
    doc.TrackRevisions = originalTracking
    Set inserted = doc.Range(startAt, target.End)
    Exit Sub
Failed:
    message = CLExplain(Err.number, Err.Description)
    On Error Resume Next
    If undoStarted Then Application.UndoRecord.EndCustomRecord
    doc.TrackRevisions = originalTracking
    If attempted Then
        doc.Activate
        If Not doc.Undo(1) Then
            message = message & " Word could not undo the part-finished insertion - check this document before continuing."
        End If
    End If
    On Error GoTo 0
    CLFail message
End Sub

' Selects the first [placeholder] or {{placeholder}} in what was just inserted,
' which is more use at that moment than a warning beforehand was.
Public Function CLSelectFirstPlaceholder(ByVal inserted As Range) As Boolean
    Dim probe As Range, pattern As Variant
    On Error Resume Next
    For Each pattern In Array("\{\{*\}\}", "\[[!\]]@\]")
        Set probe = inserted.Duplicate
        With probe.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = CStr(pattern): .MatchWildcards = True
            .Forward = True: .Wrap = wdFindStop: .Format = False
        End With
        If probe.Find.Execute Then
            If probe.Start >= inserted.Start And probe.End <= inserted.End Then
                probe.Select
                CLSelectFirstPlaceholder = True
                Exit Function
            End If
        End If
    Next
End Function

' The wording as it will look, including the numbers Word generates. Used for
' the preview panel, where a lawyer decides whether this is the right clause.
Public Function CLNumberedText(ByVal content As Range) As String
    Dim p As Paragraph, marker As String, line As String, out As String, count As Long
    On Error GoTo Plain
    For Each p In content.Paragraphs
        count = count + 1
        If count > 400 Then out = out & vbLf & "...": Exit For
        marker = ""
        If p.Range.ListFormat.ListType <> wdListNoNumbering Then marker = Trim$(p.Range.ListFormat.ListString)
        line = CLSafeText(p.Range.Text)
        Do While Right$(line, 1) = vbLf: line = Left$(line, Len(line) - 1): Loop
        If Len(marker) > 0 Then line = marker & vbTab & line
        out = out & line & vbLf
    Next
    CLNumberedText = CLTrimWording(out)
    Exit Function
Plain:
    CLNumberedText = CLPlainWording(content.Text)
End Function

' A readable description of structure, used by the self-check to prove that
' capture and insertion preserve lists and tables.
Public Function CLStructuralSignature(ByVal content As Range) As String
    Dim p As Paragraph, t As Table, result As String, r As Range
    result = "text=" & CLPlainWording(content.Text) & vbLf & "tables=" & content.Tables.count
    For Each t In content.Tables
        result = result & vbLf & "tablecells=" & t.Range.Cells.count
    Next
    For Each p In content.Paragraphs
        Set r = p.Range
        result = result & vbLf & "paragraph=" & r.ListFormat.ListType & ":" & r.ListFormat.ListLevelNumber & ":" & r.ListFormat.ListString
    Next
    CLStructuralSignature = result
End Function
