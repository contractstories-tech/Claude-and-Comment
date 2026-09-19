Attribute VB_Name = "CLSetup"
' Clause Library - setting up, and taking it away again.
'
' Setup touches three things and nothing else: it copies the template into
' Word's own Startup folder, records where your library is, and keeps a copy of
' the setup document so removal is always possible. Normal.dotm, your trust
' settings and any organisation policy are left exactly as they are.
Option Explicit

Private Const SETTINGS As String = "ClauseLibraryPersonal"
Private Const TEMPLATE As String = "ClauseLibraryPersonal.dotm"
Private Const SUPPORTDOC As String = "Setup and removal.docm"

Public Sub CLSetUp()
    On Error GoTo Failed
    Dim source As String, destination As String, root As String, startup As String, copied As Boolean
    source = ThisDocument.path & "\" & TEMPLATE
    If Not CLExists(source) Then
        CLFail "The Clause Library template was not found next to this document." & vbCrLf & vbCrLf & _
               "Extract the whole downloaded folder first - do not open this document from inside the zip file - then open it again from the extracted folder."
    End If
    startup = Options.DefaultFilePath(wdStartupPath)
    If Len(startup) = 0 Then CLFail "Word's Startup folder is not available on this computer, so nothing was changed."
    If Not CLFolderExists(startup) Then CLFolder startup
    destination = startup & "\" & TEMPLATE
    If CLExists(destination) Then
        ' A library is meant to outlast several releases, so replacing the
        ' installed template has to be an ordinary supported action rather
        ' than remove-and-hope.
        Dim installed As String: installed = GetSetting(SETTINGS, "Installation", "Version", "an earlier version")
        If MsgBox("Clause Library " & installed & " is already set up in Word." & vbCrLf & vbCrLf & _
                  "Replace it with version " & CL_VERSION & "?" & vbCrLf & vbCrLf & _
                  "Your clauses, their previous versions, your backups and the folder you are connected to are all left " & _
                  "exactly as they are. Only the program itself is replaced.", _
                  vbOKCancel + vbQuestion, "Update Clause Library") <> vbOK Then Exit Sub
        CLReplaceTemplate source, destination
        SaveSetting SETTINGS, "Installation", "Version", CL_VERSION
        SaveSetting SETTINGS, "Installation", "UpdatedOn", CLStamp()
        MsgBox "Updated to version " & CL_VERSION & "." & vbCrLf & vbCrLf & _
               "Close every Word window and open Word again to finish. Then run Options > Run self-check." & vbCrLf & vbCrLf & _
               "Your library was not touched.", vbInformation, "Update complete"
        Exit Sub
    End If

    root = Environ$("LOCALAPPDATA") & "\ClauseLibraryPersonal"
    If Len(GetSetting(SETTINGS, "Library", "Connection", "")) > 0 Then
        On Error Resume Next
        root = CLRoot()
        If Err.number <> 0 Then
            Err.Clear
            root = Environ$("LOCALAPPDATA") & "\ClauseLibraryPersonal"
        End If
        On Error GoTo Failed
    End If
    If CLFolderExists(root) Then
        If Not CLExists(root & "\library.xml") Then
            CLFail "The folder " & root & " already contains other files, so no library was created there." & vbCrLf & vbCrLf & _
                   "Nothing was overwritten. Use Use an existing library on this tab to choose a different folder."
        End If
    Else
        CLCreateLibrary root
    End If

    CLFso().CopyFile source, destination, False: copied = True
    If CLFileSha(source) <> CLFileSha(destination) Then CLFail "The copy of the template could not be verified, so setup stopped."
    CLWrite destination & ".owner", CL_PRODUCT & vbCrLf & destination
    CLRemember root
    SaveSetting SETTINGS, "Installation", "Template", destination
    SaveSetting SETTINGS, "Installation", "Version", CL_VERSION
    SaveSetting SETTINGS, "Installation", "SetUpOn", CLStamp()

    ' A copy kept outside the library folder, so that removal still works even
    ' if the library is later moved, restored or kept on a drive that is away.
    Dim supportRoot As String, supportDocument As String
    supportRoot = Environ$("LOCALAPPDATA") & "\ClauseLibrarySupport"
    If CLFolderExists(supportRoot) Then
        If Not CLExists(supportRoot & "\owner.txt") Then
            CLFail "The folder " & supportRoot & " already exists and does not belong to Clause Library. Nothing in it was changed."
        End If
        If CLRead(supportRoot & "\owner.txt") <> CL_PRODUCT Then CLFail "The support folder belongs to another application. Nothing in it was changed."
    Else
        CLFolder supportRoot
        CLWrite supportRoot & "\owner.txt", CL_PRODUCT
    End If
    supportDocument = supportRoot & "\" & SUPPORTDOC
    If LCase$(Right$(ThisDocument.FullName, 5)) <> ".docm" Then CLFail "Run setup from the Start Here document."
    If StrComp(ThisDocument.FullName, supportDocument, vbTextCompare) <> 0 Then
        CLFso().CopyFile ThisDocument.FullName, supportDocument, True
    End If
    If StrComp(source, supportRoot & "\" & TEMPLATE, vbTextCompare) <> 0 Then
        CLFso().CopyFile source, supportRoot & "\" & TEMPLATE, True
    End If
    SaveSetting SETTINGS, "Installation", "SetupDocument", supportDocument

    AddIns.Add FileName:=destination, Install:=True
    MsgBox "Clause Library is ready." & vbCrLf & vbCrLf & _
           "Look for the Clause Library tab in any Word document - it will be there after restarts too." & vbCrLf & vbCrLf & _
           "Start by selecting some wording you want to keep and choosing Capture selection." & vbCrLf & vbCrLf & _
           "Your library is stored at:" & vbCrLf & root & vbCrLf & vbCrLf & _
           "Removing Clause Library later never deletes it.", vbInformation, "Set up complete"
    Exit Sub
Failed:
    Dim message As String: message = CLExplain(Err.number, Err.Description)
    If copied Then
        On Error Resume Next
        Dim a As AddIn
        For Each a In AddIns
            If StrComp(a.path & "\" & a.Name, destination, vbTextCompare) = 0 Then a.Installed = False
        Next
        CLFso().DeleteFile destination
        If CLExists(destination & ".owner") Then CLFso().DeleteFile destination & ".owner"
        DeleteSetting SETTINGS, "Installation"
        If CLExists(destination) Then
            message = message & vbCrLf & vbCrLf & "Word is still holding the new template open. Close Word completely, then delete this file:" & vbCrLf & destination
        End If
        On Error GoTo 0
    End If
    MsgBox "Setup did not finish." & vbCrLf & vbCrLf & message & vbCrLf & vbCrLf & "Your saved clauses were not affected.", _
           vbExclamation, "Clause Library"
End Sub

' Swaps the installed template for a new one. The old file is kept until the
' new one is verified, so a failed update leaves a working installation.
Private Sub CLReplaceTemplate(ByVal source As String, ByVal destination As String)
    Dim keep As String, a As AddIn
    keep = destination & ".previous"
    On Error Resume Next
    For Each a In AddIns
        If StrComp(a.path & "\" & a.Name, destination, vbTextCompare) = 0 Then a.Installed = False
    Next
    If CLExists(keep) Then CLFso().DeleteFile keep
    On Error GoTo Failed
    CLFso().MoveFile destination, keep
    CLFso().CopyFile source, destination, True
    If CLFileSha(source) <> CLFileSha(destination) Then CLFail "The new template could not be verified."
    CLWrite destination & ".owner", CL_PRODUCT & vbCrLf & destination
    On Error Resume Next
    CLFso().DeleteFile keep
    AddIns.Add FileName:=destination, Install:=True
    On Error GoTo 0
    Exit Sub
Failed:
    Dim message As String: message = CLExplain(Err.number, Err.Description)
    On Error Resume Next
    If Not CLExists(destination) And CLExists(keep) Then CLFso().MoveFile keep, destination
    AddIns.Add FileName:=destination, Install:=True
    On Error GoTo 0
    CLFail "The update did not complete, and the version you had is still in place. " & message & vbCrLf & vbCrLf & _
           "If Word is holding the template open, close every Word window and try again."
End Sub

' Takes the tools out of Word and leaves every clause where it is.
Public Sub CLRemoveIntegration()
    On Error GoTo Failed
    Dim destination As String, startupCopy As String
    destination = GetSetting(SETTINGS, "Installation", "Template", "")
    startupCopy = Options.DefaultFilePath(wdStartupPath) & "\" & TEMPLATE
    If Len(destination) = 0 Then destination = startupCopy
    If Not CLExists(destination) And CLExists(startupCopy) Then destination = startupCopy
    If Not CLExists(destination) Then
        CLFail "Clause Library does not appear to be set up in Word on this computer. Nothing was changed."
    End If
    If StrComp(destination, startupCopy, vbTextCompare) <> 0 Then
        CLFail "The recorded template location is not Word's Startup folder, so no file was removed." & vbCrLf & vbCrLf & _
               "Expected: " & startupCopy & vbCrLf & "Recorded: " & destination
    End If
    If CLExists(destination & ".owner") Then
        If CLRead(destination & ".owner") <> CL_PRODUCT & vbCrLf & destination Then
            CLFail "The ownership record next to that template does not match, so no file was removed."
        End If
    End If
    If MsgBox("Remove the Clause Library tab and tools from Word?" & vbCrLf & vbCrLf & _
              "Your clauses, previous versions and backups all stay exactly where they are. " & _
              "You can set it up again at any time and reconnect to the same library.", _
              vbOKCancel + vbQuestion, "Remove Word integration") <> vbOK Then Exit Sub
    CLReleaseManager
    Dim a As AddIn
    For Each a In AddIns
        If StrComp(a.path & "\" & a.Name, destination, vbTextCompare) = 0 Then a.Installed = False
    Next
    If CLExists(destination) Then CLFso().DeleteFile destination
    If CLExists(destination & ".owner") Then CLFso().DeleteFile destination & ".owner"
    DeleteSetting SETTINGS, "Installation"
    MsgBox "Removed from Word." & vbCrLf & vbCrLf & _
           "Your library and its location were kept. Close and reopen Word to clear the tab from any window that is still showing it.", _
           vbInformation, "Clause Library"
    Exit Sub
Failed:
    MsgBox "Removal did not finish." & vbCrLf & vbCrLf & CLExplain(Err.number, Err.Description) & vbCrLf & vbCrLf & _
           "No clauses were deleted. If Word is holding the template open, close every Word window and try again from Setup and removal.docm.", _
           vbExclamation, "Clause Library"
End Sub

' The complete removal a person needs before handing a computer back. It says
' exactly what it will delete and requires the word DELETE to be typed.
Public Sub CLRemoveEverything()
    On Error GoTo Failed
    Dim root As String, supportRoot As String, items As String
    supportRoot = Environ$("LOCALAPPDATA") & "\ClauseLibrarySupport"
    On Error Resume Next
    root = CLRoot()
    On Error GoTo Failed
    items = "  The Clause Library tab and template in Word" & vbCrLf
    If Len(root) > 0 Then items = items & "  Your whole library, including every clause, private note and previous version:" & vbCrLf & "      " & root & vbCrLf
    If CLFolderExists(supportRoot) Then items = items & "  The setup copy kept at " & supportRoot & vbCrLf
    items = items & "  The saved library location in your Windows profile" & vbCrLf

    If MsgBox("This deletes Clause Library AND YOUR CLAUSES from this computer:" & vbCrLf & vbCrLf & items & vbCrLf & _
              "Backups you saved elsewhere are NOT touched." & vbCrLf & vbCrLf & _
              "This cannot be undone. Continue?", vbYesNo + vbCritical + vbDefaultButton2, "Remove everything") <> vbYes Then Exit Sub
    Dim typed As String
    typed = InputBox("To confirm, type the word DELETE and press OK." & vbCrLf & vbCrLf & _
                     "Anything else cancels and nothing will be removed.", "Remove everything")
    If UCase$(Trim$(typed)) <> "DELETE" Then
        MsgBox "Nothing was removed.", vbInformation, "Cancelled"
        Exit Sub
    End If

    CLReleaseManager
    Dim removed As String
    On Error Resume Next
    Dim destination As String: destination = Options.DefaultFilePath(wdStartupPath) & "\" & TEMPLATE
    Dim a As AddIn
    For Each a In AddIns
        If StrComp(a.path & "\" & a.Name, destination, vbTextCompare) = 0 Then a.Installed = False
    Next
    If CLExists(destination) Then CLFso().DeleteFile destination: removed = removed & "  Word template" & vbCrLf
    If CLExists(destination & ".owner") Then CLFso().DeleteFile destination & ".owner"
    If Len(root) > 0 And CLFolderExists(root) Then
        CLFso().DeleteFolder root, True
        removed = removed & "  Library folder" & vbCrLf
    End If
    If CLFolderExists(supportRoot) Then
        CLFso().DeleteFolder supportRoot, True
        removed = removed & "  Support folder" & vbCrLf
    End If
    DeleteSetting SETTINGS
    removed = removed & "  Saved settings" & vbCrLf
    On Error GoTo 0
    MsgBox "Clause Library has been removed from this computer." & vbCrLf & vbCrLf & removed & vbCrLf & _
           "Close and reopen Word to clear the tab. Any backups you saved elsewhere are still there.", _
           vbInformation, "Removed"
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbExclamation, "Removal did not finish"
End Sub

Public Sub CLOpenRemoval()
    On Error GoTo Failed
    If Not CLCloseManager() Then Exit Sub
    Dim setupDocument As String: setupDocument = GetSetting(SETTINGS, "Installation", "SetupDocument", "")
    If Len(setupDocument) = 0 Or Not CLExists(setupDocument) Then
        CLFail "The saved setup document could not be found. Open the Start Here document from the folder you extracted, and use Remove Word integration there."
    End If
    CLRequireLocalPath setupDocument, "The saved setup document"
    Documents.Open FileName:=setupDocument
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbInformation, "Clause Library"
End Sub

Public Sub CLSetupLocate()
    On Error GoTo Failed
    Dim picker As Object: Set picker = Application.FileDialog(4)
    picker.Title = "Choose your existing Clause Library folder"
    If picker.Show <> -1 Then Exit Sub
    CLRemember CStr(picker.SelectedItems(1))
    MsgBox "Connected to:" & vbCrLf & vbCrLf & picker.SelectedItems(1) & vbCrLf & vbCrLf & _
           "Now choose Set up Clause Library to make it available in Word.", vbInformation, "Clause Library"
    Exit Sub
Failed:
    MsgBox CLExplain(Err.number, Err.Description), vbInformation, "That library was not connected"
End Sub

Public Sub CLRibbonSetup(ByVal control As Object): CLSetUp: End Sub
Public Sub CLRibbonRemove(ByVal control As Object): CLOpenRemoval: End Sub
Public Sub CLRibbonRemoveHere(ByVal control As Object): CLRemoveIntegration: End Sub
Public Sub CLRibbonRemoveEverything(ByVal control As Object): CLRemoveEverything: End Sub
Public Sub CLRibbonSetupLocate(ByVal control As Object): CLSetupLocate: End Sub
