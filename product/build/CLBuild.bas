Attribute VB_Name = "CLBuild"
' Clause Library - the builder.
'
' Run this once, on a Windows computer with desktop Word, to turn the source in
' word\ into the two finished files you hand to people. Word writes the macro
' project itself, which is the only way to be certain it is correct.
'
' HOW TO RUN IT
'   1. In Word: File > Options > Trust Center > Trust Center Settings >
'      Macro Settings, and tick "Trust access to the VBA project object model".
'   2. Open a NEW, EMPTY Word document. Press Alt+F11.
'   3. File > Import File, and choose this CLBuild.bas.
'   4. Put the cursor inside BuildClauseLibrary below and press F5.
'   5. Untick the Trust Center setting again afterwards. Nobody you give the
'      finished files to ever needs it.
'
' It writes everything into a "dist" folder beside the product folder and
' leaves your own library and settings completely alone.
Option Explicit

Private Const PROJECT_RUNTIME As String = "ClauseLibraryProduct"
Private Const PROJECT_SETUP As String = "ClauseLibrarySetup"

Private root As String
Private report As String

Public Sub BuildClauseLibrary()
    Dim fso As Object, dist As String, started As Date
    started = Now
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")

    root = FindRoot(fso)
    If Len(root) = 0 Then
        MsgBox "Could not find the Clause Library source." & vbCrLf & vbCrLf & _
               "Put this CLBuild.bas inside the product\build folder, next to the word and seed folders, " & _
               "then import it again from there." & vbCrLf & vbCrLf & _
               "Looked beside: " & ThisDocument.path, vbExclamation, "Build"
        Exit Sub
    End If
    If Not CanReachVbaProject() Then
        MsgBox "Word is not allowing this macro to build a project." & vbCrLf & vbCrLf & _
               "Turn on: File > Options > Trust Center > Trust Center Settings > Macro Settings >" & vbCrLf & _
               "         Trust access to the VBA project object model" & vbCrLf & vbCrLf & _
               "Then run this again. Turn it back off when the build has finished - " & _
               "the finished files never need it.", vbExclamation, "Build"
        Exit Sub
    End If

    report = "CLAUSE LIBRARY BUILD" & vbCrLf & "Started " & Now & vbCrLf & _
             "Word " & Application.Version & " " & Application.Build & vbCrLf & _
             "Source: " & root & vbCrLf & String$(68, "=") & vbCrLf

    ValidateSources fso

    dist = fso.GetParentFolderName(root) & "\dist"
    If fso.FolderExists(dist) Then fso.DeleteFolder dist, True
    fso.CreateFolder dist

    Say "Building the Word template..."
    BuildRuntime fso, dist
    Say "Building the setup document..."
    BuildSetup fso, dist
    Say "Copying the documents..."
    CopyDocs fso, dist
    WriteManifest fso, dist

    report = report & String$(68, "=") & vbCrLf & "Finished in " & DateDiff("s", started, Now) & " seconds." & vbCrLf
    fso.CreateTextFile(dist & "\build-log.txt", True).Write report
    Application.StatusBar = False
    MsgBox "Build finished." & vbCrLf & vbCrLf & _
           "The folder to give people is:" & vbCrLf & dist & vbCrLf & vbCrLf & _
           "Before you send it to anyone:" & vbCrLf & _
           "  1. Open Start Here.docm from that folder and set it up." & vbCrLf & _
           "  2. On the Clause Library tab choose Options > Run self-check." & vbCrLf & _
           "  3. Only send it on if every check passes." & vbCrLf & vbCrLf & _
           "Remember to turn off 'Trust access to the VBA project object model' again.", _
           vbInformation, "Build complete"
    Exit Sub
Failed:
    Application.StatusBar = False
    MsgBox "The build stopped." & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & report, vbCritical, "Build"
End Sub

Private Sub Say(ByVal message As String)
    Application.StatusBar = "Clause Library build: " & message
    report = report & message & vbCrLf
End Sub

Private Function FindRoot(ByVal fso As Object) As String
    Dim here As String, candidates As Variant, c As Variant
    here = ThisDocument.path
    candidates = Array(here, fso.GetParentFolderName(here), here & "\product", _
                       fso.GetParentFolderName(here) & "\product")
    For Each c In candidates
        If Len(c) > 0 Then
            If fso.FolderExists(c & "\word") And fso.FolderExists(c & "\build\seed") Then FindRoot = c: Exit Function
        End If
    Next
End Function

Private Function CanReachVbaProject() As Boolean
    On Error Resume Next
    Dim n As Long
    n = ThisDocument.VBProject.VBComponents.count
    CanReachVbaProject = (Err.number = 0)
    Err.Clear
End Function

' ---------- checking the source before using it ----------

Private Sub ValidateSources(ByVal fso As Object)
    Dim f As Object, names As Object, text As String, lines_ As Variant, i As Long
    Dim line As String, nameOf As String, problems As String, files As Long
    Set names = CreateObject("Scripting.Dictionary")
    For Each f In fso.GetFolder(root & "\word").Files
        If LCase$(Right$(f.Name, 4)) = ".bas" Or LCase$(Right$(f.Name, 4)) = ".txt" Then
            files = files + 1
            text = fso.OpenTextFile(f.path, 1).ReadAll
            lines_ = Split(Replace$(text, vbCrLf, vbLf), vbLf)
            For i = 0 To UBound(lines_)
                line = CStr(lines_(i))
                If Len(line) > 1023 Then problems = problems & "  " & f.Name & " line " & (i + 1) & " is too long for VBA" & vbCrLf
                nameOf = PublicProcedureName(line)
                If Len(nameOf) > 0 And LCase$(Right$(f.Name, 4)) = ".bas" Then
                    If names.Exists(LCase$(nameOf)) Then
                        problems = problems & "  " & nameOf & " is declared in both " & f.Name & " and " & names(LCase$(nameOf)) & vbCrLf
                    Else
                        names.Add LCase$(nameOf), f.Name
                    End If
                End If
            Next
        End If
    Next
    If Len(problems) > 0 Then Err.Raise 5, , "The source has problems that would not compile:" & vbCrLf & problems
    report = report & files & " source files checked: line lengths and duplicate public names are clean." & vbCrLf
End Sub

Private Function PublicProcedureName(ByVal line As String) As String
    Dim t As String: t = LTrim$(line)
    If LCase$(Left$(t, 7)) = "private" Then Exit Function
    If LCase$(Left$(t, 6)) = "public" Then t = LTrim$(Mid$(t, 7))
    If LCase$(Left$(t, 7)) = "static " Then t = LTrim$(Mid$(t, 8))
    Dim kind As String
    If LCase$(Left$(t, 4)) = "sub " Then
        t = Mid$(t, 5)
    ElseIf LCase$(Left$(t, 9)) = "function " Then
        t = Mid$(t, 10)
    Else
        Exit Function
    End If
    Dim i As Long, ch As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If Not (ch Like "[A-Za-z0-9_]") Then Exit For
    Next
    PublicProcedureName = Left$(t, i - 1)
End Function

' ---------- the runtime template ----------

Private Sub BuildRuntime(ByVal fso As Object, ByVal dist As String)
    Dim working As String, d As Document
    working = dist & "\ClauseLibraryPersonal.dotm"
    fso.CopyFile root & "\build\seed\ClauseLibraryRuntime.dotm", working, True
    Set d = Documents.Open(FileName:=working, AddToRecentFiles:=False, Visible:=False)
    ClearProject d
    d.VBProject.Name = PROJECT_RUNTIME
    ImportModules fso, d, Array("CLPlatform", "CLStore", "CLRich", "CLActions", _
                                "CLExport", "CLRecovery", "CLSetup", "CLSelfCheck")
    BuildLibraryForm fso, d
    BuildHistoryForm fso, d
    d.Save
    d.Close wdDoNotSaveChanges
    report = report & "  ClauseLibraryPersonal.dotm: 8 modules and 2 windows" & vbCrLf
End Sub

Private Sub BuildSetup(ByVal fso As Object, ByVal dist As String)
    Dim working As String, d As Document, shim As Object
    working = dist & "\Start Here.docm"
    fso.CopyFile root & "\build\seed\Start Here.docm", working, True
    Set d = Documents.Open(FileName:=working, AddToRecentFiles:=False, Visible:=False)
    ClearProject d
    d.VBProject.Name = PROJECT_SETUP
    ImportModules fso, d, Array("CLPlatform", "CLStore", "CLSetup")
    ' The setup document has no library window, so it supplies the two calls
    ' CLStore expects rather than carrying the whole of CLActions.
    Set shim = d.VBProject.VBComponents.Add(1)
    shim.Name = "SetupSupport"
    shim.CodeModule.AddFromString _
        "Option Explicit" & vbCrLf & _
        "' Stand-ins for the library window, which this document does not have." & vbCrLf & _
        "Public Function CLCloseManager() As Boolean" & vbCrLf & _
        "    CLCloseManager = True" & vbCrLf & _
        "End Function" & vbCrLf & _
        "Public Sub CLReleaseManager()" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Public Function CLEnsureReady() As Boolean" & vbCrLf & _
        "    CLEnsureReady = True" & vbCrLf & _
        "End Function" & vbCrLf & _
        "Public Sub CLNotifyManager()" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Public Sub CLRefreshRibbon()" & vbCrLf & _
        "End Sub"
    d.Save
    d.Close wdDoNotSaveChanges
    report = report & "  Start Here.docm: 3 modules and the setup shim" & vbCrLf
End Sub

Private Sub ClearProject(ByVal d As Document)
    Dim i As Long, comp As Object
    For i = d.VBProject.VBComponents.count To 1 Step -1
        Set comp = d.VBProject.VBComponents(i)
        If comp.Type = 100 Then                       ' ThisDocument: cannot be removed, so empty it
            If comp.CodeModule.CountOfLines > 0 Then comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
        Else
            d.VBProject.VBComponents.Remove comp
        End If
    Next
End Sub

Private Sub ImportModules(ByVal fso As Object, ByVal d As Document, ByVal names As Variant)
    Dim n As Variant, path As String
    For Each n In names
        path = root & "\word\" & n & ".bas"
        If Not fso.FileExists(path) Then Err.Raise 5, , "Missing source file: " & path
        d.VBProject.VBComponents.Import path
    Next
End Sub

Private Function ReadCode(ByVal fso As Object, ByVal path As String) As String
    If Not fso.FileExists(path) Then Err.Raise 5, , "Missing source file: " & path
    ReadCode = fso.OpenTextFile(path, 1).ReadAll
End Function

' ---------- the two windows ----------

' Controls are created here; where they sit is decided at run time by the
' form's own LayOut, so the window fits the screen it is actually on.
Private Sub BuildLibraryForm(ByVal fso As Object, ByVal d As Document)
    Dim comp As Object, designer As Object
    Set comp = d.VBProject.VBComponents.Add(3)
    comp.Name = "frmLibrary"
    Set designer = comp.designer
    On Error Resume Next
    comp.Properties("Width") = 960
    comp.Properties("Height") = 600
    On Error GoTo 0

    Add designer, "Label", "lblHeading", 10, 10, 300, 20
    Add designer, "Label", "lblSearch", 10, 34, 300, 14
    Add designer, "TextBox", "txtSearch", 10, 50, 300, 24
    Add designer, "ComboBox", "cboView", 10, 80, 140, 22
    Add designer, "ComboBox", "cboSort", 155, 80, 100, 22
    Add designer, "CommandButton", "cmdRefresh", 260, 80, 50, 22
    Add designer, "Label", "lblCount", 10, 106, 300, 14
    Add designer, "Label", "lblC1", 10, 122, 120, 13
    Add designer, "Label", "lblC2", 130, 122, 90, 13
    Add designer, "Label", "lblC3", 220, 122, 50, 13
    Add designer, "Label", "lblC4", 270, 122, 50, 13
    Add designer, "ListBox", "lstEntries", 10, 138, 300, 300
    Add designer, "CommandButton", "cmdExport", 10, 442, 95, 24
    Add designer, "CommandButton", "cmdFolder", 110, 442, 95, 24
    Add designer, "CommandButton", "cmdHelp", 210, 442, 95, 24

    Add designer, "Label", "lblTitle", 330, 10, 300, 14
    Add designer, "TextBox", "txtTitle", 330, 26, 300, 22
    Add designer, "Label", "lblTopic", 330, 54, 145, 14
    Add designer, "TextBox", "txtTopic", 330, 69, 145, 21
    Add designer, "Label", "lblTags", 485, 54, 145, 14
    Add designer, "TextBox", "txtTags", 485, 69, 145, 21
    Add designer, "Label", "lblFamily", 330, 95, 145, 14
    Add designer, "TextBox", "txtFamily", 330, 110, 145, 21
    Add designer, "Label", "lblRole", 485, 95, 145, 14
    Add designer, "ComboBox", "cboRole", 485, 110, 145, 21
    Add designer, "Label", "lblUse", 330, 136, 300, 14
    Add designer, "TextBox", "txtUseWhen", 330, 151, 300, 32
    Add designer, "Label", "lblNotes", 330, 188, 300, 14
    Add designer, "TextBox", "txtNotes", 330, 203, 300, 38
    Add designer, "CheckBox", "chkFavourite", 330, 246, 90, 18
    Add designer, "CheckBox", "chkOrganised", 425, 246, 100, 18
    Add designer, "CommandButton", "cmdSave", 530, 244, 100, 24
    Add designer, "Label", "lblPreview", 330, 272, 300, 14
    Add designer, "TextBox", "txtPreview", 330, 288, 300, 130
    Add designer, "CommandButton", "cmdInsert", 330, 424, 95, 26
    Add designer, "CommandButton", "cmdPlain", 430, 424, 95, 26
    Add designer, "CommandButton", "cmdEdit", 530, 424, 100, 26
    Add designer, "CommandButton", "cmdArchive", 330, 454, 58, 22
    Add designer, "CommandButton", "cmdTrash", 392, 454, 58, 22
    Add designer, "CommandButton", "cmdRestore", 454, 454, 58, 22
    Add designer, "CommandButton", "cmdHistory", 516, 454, 58, 22
    Add designer, "CommandButton", "cmdClose", 578, 454, 52, 22
    Add designer, "Label", "lblStatus", 10, 480, 620, 15

    comp.CodeModule.AddFromString ReadCode(fso, root & "\word\LibraryForm.txt")
End Sub

Private Sub BuildHistoryForm(ByVal fso As Object, ByVal d As Document)
    Dim comp As Object, designer As Object
    Set comp = d.VBProject.VBComponents.Add(3)
    comp.Name = "frmHistory"
    Set designer = comp.designer
    On Error Resume Next
    comp.Properties("Width") = 720
    comp.Properties("Height") = 470
    On Error GoTo 0
    Add designer, "Label", "lblExplain", 12, 10, 690, 28
    Add designer, "ListBox", "lstVersions", 12, 44, 690, 120
    Add designer, "TextBox", "txtHistory", 12, 172, 690, 200
    Add designer, "CommandButton", "cmdCompare", 12, 382, 150, 26
    Add designer, "CommandButton", "cmdRecoverWording", 170, 382, 168, 26
    Add designer, "CommandButton", "cmdRecover", 346, 382, 168, 26
    Add designer, "CommandButton", "cmdDone", 592, 382, 110, 26
    comp.CodeModule.AddFromString ReadCode(fso, root & "\word\HistoryForm.txt")
End Sub

Private Sub Add(ByVal designer As Object, ByVal kind As String, ByVal name As String, _
                ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = designer.Controls.Add("Forms." & kind & ".1", name, True)
    c.Left = x: c.Top = y: c.Width = w: c.Height = h
End Sub

' ---------- documents and checksums ----------

Private Sub CopyDocs(ByVal fso As Object, ByVal dist As String)
    Dim n As Variant
    For Each n In Array("READ ME FIRST.txt", "What changed in 2.0.txt")
        If fso.FileExists(root & "\docs\" & n) Then fso.CopyFile root & "\docs\" & n, dist & "\" & n, True
    Next
End Sub

Private Sub WriteManifest(ByVal fso As Object, ByVal dist As String)
    Dim f As Object, text As String
    For Each f In fso.GetFolder(dist).Files
        If LCase$(f.Name) <> "manifest_sha256.txt" And LCase$(f.Name) <> "build-log.txt" Then
            text = text & FileSha(f.path) & "  " & f.Name & vbCrLf
        End If
    Next
    fso.CreateTextFile(dist & "\MANIFEST_SHA256.txt", True).Write text
    report = report & vbCrLf & "Checksums of what will be shipped:" & vbCrLf & text
End Sub

' Uses Windows' own certutil rather than the product's checksum code, so the
' manifest is not produced by the thing it is meant to vouch for.
Private Function FileSha(ByVal path As String) As String
    Dim shell As Object, fso As Object, temp As String, text As String, lines_ As Variant, i As Long
    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    temp = fso.GetSpecialFolder(2) & "\clbuild-" & Format$(Now, "hhnnss") & "-" & Int(Rnd * 10000) & ".txt"
    shell.Run "cmd /c certutil -hashfile """ & path & """ SHA256 > """ & temp & """", 0, True
    If Not fso.FileExists(temp) Then FileSha = "(checksum unavailable)": Exit Function
    text = fso.OpenTextFile(temp, 1).ReadAll
    fso.DeleteFile temp
    lines_ = Split(Replace$(text, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines_)
        If Len(Replace$(Trim$(CStr(lines_(i))), " ", "")) = 64 Then
            If Not (Trim$(CStr(lines_(i))) Like "*[!0-9a-fA-F ]*") Then
                FileSha = LCase$(Replace$(Trim$(CStr(lines_(i))), " ", ""))
                Exit Function
            End If
        End If
    Next
    FileSha = "(checksum unavailable)"
End Function
