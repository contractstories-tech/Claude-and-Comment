Attribute VB_Name = "CLExport"
' Clause Library - getting your wording out.
'
' A precedent collection you cannot read without a particular tool is not
' really yours. Export writes ordinary Word files, a plain index you can open
' in any browser, and a spreadsheet of the details - no macros, no library, no
' Clause Library needed to read any of it.
Option Explicit

Public Sub CLExportLibrary()
    On Error GoTo Failed
    If Not CLEnsureReady() Then Exit Sub
    Dim rows As Collection, active As Collection, r As Variant
    Set rows = CLList()
    Set active = New Collection
    For Each r In rows
        If r("state") = "Active" Then active.Add r
    Next
    If active.count = 0 Then
        MsgBox "There is nothing to export yet. Capture some wording first.", vbInformation, "Export library"
        Exit Sub
    End If

    Dim picker As Object: Set picker = Application.FileDialog(4)
    picker.Title = "Choose a folder to export your clause library into"
    If picker.Show <> -1 Then Exit Sub

    CLRequireLocalPath CStr(picker.SelectedItems(1)), "The folder you chose"
    If Not CLConfirmSyncedFolder(CStr(picker.SelectedItems(1)), "Every clause you export") Then Exit Sub

    Dim includeNotes As VbMsgBoxResult
    includeNotes = MsgBox("Include your private notes in the exported details file?" & vbCrLf & vbCrLf & _
                          "Choose No if you intend to share this export with anyone.", _
                          vbYesNoCancel + vbQuestion + vbDefaultButton2, "Export library")
    If includeNotes = vbCancel Then Exit Sub

    If MsgBox("Export " & active.count & " items as Word files, plus an index and a details file?" & vbCrLf & vbCrLf & _
              "This can take a little while for a large library. Word will be busy until it finishes.", _
              vbOKCancel + vbInformation, "Export library") <> vbOK Then Exit Sub

    Dim folder As String, clauses As String, written As Long, failed As String
    folder = picker.SelectedItems(1) & "\Clause Library export " & Format$(Now, "yyyy-mm-dd")
    If CLFolderExists(folder) Then folder = folder & " " & Format$(Now, "hhnnss")
    clauses = folder & "\Clauses"
    CLFolder folder: CLFolder clauses

    Dim csv As String, html As String, index As Long, name As String
    csv = "Number,Title,Kind,Topic,Tags,Family,Variant role,Use when,Captured,Last changed,Times used,From document,File"
    If includeNotes = vbYes Then csv = csv & ",Private notes"
    csv = csv & vbCrLf
    html = CLIndexHead(active.count)

    Application.ScreenUpdating = False
    For Each r In active
        index = index + 1
        Application.StatusBar = "Exporting " & index & " of " & active.count & "..."
        name = Format$(index, "000") & " " & CLSafeFileName(r("title"), 70)
        On Error Resume Next
        Err.Clear
        CLWriteEntryDocument CStr(r("id")), clauses & "\" & name & ".docx"
        If Err.number <> 0 Then
            failed = failed & "  " & r("title") & vbCrLf
            Err.Clear
        Else
            written = written + 1
        End If
        On Error GoTo Failed
        csv = csv & index & "," & CLCsv(r("title")) & "," & CLCsv(r("kind")) & "," & CLCsv(r("topic")) & "," & _
              CLCsv(r("tags")) & "," & CLCsv(r("family")) & "," & CLCsv(r("role")) & "," & CLCsv(r("applicability")) & "," & _
              CLCsv(r("created")) & "," & CLCsv(r("updated")) & "," & CLCsv(r("usedCount")) & "," & _
              CLCsv(r("source")) & "," & CLCsv("Clauses\" & name & ".docx")
        If includeNotes = vbYes Then csv = csv & "," & CLCsv(r("notes"))
        csv = csv & vbCrLf
        html = html & CLIndexItem(r, "Clauses/" & CLUrl(name) & ".docx")
    Next
    Application.ScreenUpdating = True
    Application.StatusBar = False

    html = html & CLIndexFoot()
    CLWrite folder & "\Clause Library index.html", html
    CLWrite folder & "\Clause Library details.csv", csv
    CLWrite folder & "\READ ME.txt", CLExportReadme(written, includeNotes = vbYes)

    Dim message As String
    message = written & " of " & active.count & " items exported to:" & vbCrLf & folder
    If Len(failed) > 0 Then message = message & vbCrLf & vbCrLf & "These could not be written as Word files:" & vbCrLf & failed
    If includeNotes = vbYes Then message = message & vbCrLf & vbCrLf & "The details file contains your private notes."
    MsgBox message, vbInformation, "Export complete"
    On Error Resume Next
    If CLIsLocalPath(folder) Then Shell "explorer.exe " & Chr$(34) & folder & Chr$(34), vbNormalFocus
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox CLExplain(Err.number, Err.Description), vbExclamation, "The export did not finish"
End Sub

Private Sub CLWriteEntryDocument(ByVal id As String, ByVal path As String)
    Dim e As Object, d As Document, payload As String
    Set e = CLLoad(id)
    payload = CLReadPayload(e)
    Set d = Documents.Add(Visible:=False)
    d.TrackRevisions = False
    On Error Resume Next
    If Len(payload) > 0 Then d.Content.InsertXML payload
    On Error GoTo 0
    If Len(CLSafeText(d.Content.Text)) < 2 Then d.Content.Text = Replace$(CLGet(e, "plain"), vbLf, vbCr)
    d.BuiltInDocumentProperties("Title") = CLGet(e, "title")
    d.SaveAs2 FileName:=path, FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    d.Close wdDoNotSaveChanges
End Sub

Private Function CLUrl(ByVal name As String) As String
    CLUrl = Replace$(CLHtml(name), " ", "%20")
End Function

Private Function CLIndexHead(ByVal count As Long) As String
    Dim h As String
    h = "<!doctype html><html lang='en'><head><meta charset='utf-8'>" & _
        "<meta name='viewport' content='width=device-width,initial-scale=1'>" & _
        "<meta http-equiv='Content-Security-Policy' content=""default-src 'none';style-src 'unsafe-inline';script-src 'unsafe-inline'"">" & _
        "<title>Clause Library</title><style>" & _
        ":root{color-scheme:light dark}" & _
        "body{font:16px/1.6 system-ui,Segoe UI,sans-serif;margin:0;background:#f6f7f5;color:#1f2d27}" & _
        "main{max-width:60rem;margin:0 auto;padding:2rem 1rem 4rem}" & _
        "h1{font:600 2rem/1.2 Georgia,serif;margin:0 0 .25rem}" & _
        "p.lead{color:#5b6b62;margin:0 0 1.5rem}" & _
        "input{box-sizing:border-box;width:100%;padding:.75rem 1rem;border:1px solid #c3cdc6;border-radius:.5rem;font:inherit;background:#fff;color:inherit}" & _
        "article{background:#fff;border:1px solid #dde3de;border-radius:.6rem;padding:1.25rem 1.5rem;margin:1rem 0}" & _
        "h2{font:600 1.2rem/1.3 Georgia,serif;margin:0 0 .4rem}" & _
        ".meta{color:#5b6b62;font-size:.85rem;margin:0 0 .75rem}" & _
        ".use{font-style:italic;color:#3d5a4c;margin:0 0 .6rem}" & _
        "pre{white-space:pre-wrap;overflow-wrap:anywhere;font:1rem/1.6 Georgia,serif;margin:0 0 .75rem}" & _
        "a{color:#1c5c45}" & _
        "#status{min-height:1.5rem;color:#5b6b62}" & _
        "[hidden]{display:none}" & _
        "@media(prefers-color-scheme:dark){body{background:#121815;color:#e4ece7}" & _
        "article,input{background:#1b2420;border-color:#31403a;color:#e4ece7}" & _
        ".meta,.lead,#status{color:#9db0a6}a{color:#7fd0ac}}" & _
        "</style></head><body><main><h1>Clause Library</h1>" & _
        "<p class='lead'>" & count & " items, exported " & CLHtml(CLNiceDate(CLStamp())) & _
        ". Open the Word file beside each item to keep its formatting. Private notes are not shown here.</p>" & _
        "<label for='q'>Search</label><input id='q' type='search' placeholder='Try a topic, a phrase or a tag' autofocus>" & _
        "<p id='status' role='status'></p>"
    CLIndexHead = h
End Function

Private Function CLIndexItem(ByVal r As Object, ByVal href As String) As String
    Dim meta As String, s As String
    meta = r("kind")
    If Len(r("topic")) > 0 Then meta = meta & " &middot; " & r("topic")
    If Len(r("tags")) > 0 Then meta = meta & " &middot; " & r("tags")
    If Len(r("family")) > 0 Then meta = meta & " &middot; family: " & r("family")
    meta = meta & " &middot; saved " & CLNiceDate(r("created"))
    s = "<article><h2>" & CLHtml(r("title")) & "</h2><p class='meta'>" & CLHtml(meta) & "</p>"
    If Len(r("applicability")) > 0 Then s = s & "<p class='use'>" & CLHtml(r("applicability")) & "</p>"
    s = s & "<pre>" & CLHtml(r("plain")) & "</pre>" & _
        "<p><a href='" & href & "'>Open the formatted Word version</a></p></article>"
    CLIndexItem = s
End Function

Private Function CLIndexFoot() As String
    CLIndexFoot = "<script>" & _
      "var q=document.getElementById('q'),s=document.getElementById('status')," & _
      "items=Array.prototype.slice.call(document.querySelectorAll('article'));" & _
      "function run(){var v=q.value.toLowerCase();" & _
      "var w=(v.match(/""[^""]+""|[^\s]+/g)||[]).map(function(t){return t.replace(/""/g,'')}).filter(Boolean);" & _
      "var n=0;items.forEach(function(a){var t=a.textContent.toLowerCase();" & _
      "var ok=w.every(function(x){return t.indexOf(x)>-1});a.hidden=!ok;if(ok)n++;});" & _
      "s.textContent=n+(n===1?' item':' items')+(w.length&&!n?' - try fewer words':'');}" & _
      "q.addEventListener('input',run);run();</script></main></body></html>"
End Function

Private Function CLExportReadme(ByVal count As Long, ByVal withNotes As Boolean) As String
    CLExportReadme = _
      "CLAUSE LIBRARY EXPORT" & vbCrLf & vbCrLf & _
      "Exported " & CLNiceDate(CLStamp()) & " by Clause Library " & CL_VERSION & "." & vbCrLf & vbCrLf & _
      "Clauses\      one Word file per item, with its formatting" & vbCrLf & _
      "Clause Library index.html   open in any browser to search and read" & vbCrLf & _
      "Clause Library details.csv  open in Excel for topics, tags and dates" & vbCrLf & vbCrLf & _
      count & " items were exported." & vbCrLf & vbCrLf & _
      IIf(withNotes, "THE DETAILS FILE CONTAINS PRIVATE NOTES. Treat this folder as confidential." & vbCrLf & vbCrLf, _
                     "Private notes were not included." & vbCrLf & vbCrLf) & _
      "Nothing here needs Clause Library, macros or any other software to read." & vbCrLf & _
      "This is a snapshot. It does not update when your library changes." & vbCrLf
End Function

Public Sub CLRibbonExport(ByVal control As Object): CLExportLibrary: End Sub
