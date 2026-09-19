VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLDraftWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clause Library - makes a wording draft behave the way a person expects.
'
' Without this, a draft opened by Edit wording is just an ordinary Word
' document: pressing Ctrl+S writes a .docx somewhere and the library is
' untouched, while the person believes they have saved their clause. Closing it
' brings Word's generic save prompt, which means the opposite of what it says.
'
' So: in a library draft, Save means save to the library, and closing asks the
' question that actually matters.
Option Explicit

Private WithEvents app As Word.Application
Attribute app.VB_VarHelpID = -1

Public Sub Watch()
    Set app = Word.Application
End Sub

Public Sub Unwatch()
    Set app = Nothing
End Sub

Private Function IsLibraryDraft(ByVal doc As Document) As Boolean
    On Error Resume Next
    IsLibraryDraft = (Len(doc.Variables("ClauseLibraryEntry").Value) > 0)
    On Error GoTo 0
End Function

' Ctrl+S, the Save button, and every other ordinary save route.
Private Sub app_DocumentBeforeSave(ByVal doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
    If Not IsLibraryDraft(doc) Then Exit Sub
    Cancel = True
    SaveAsUI = False
    CLSaveWordingFor doc
End Sub

Private Sub app_DocumentBeforeClose(ByVal doc As Document, Cancel As Boolean)
    If Not IsLibraryDraft(doc) Then Exit Sub
    If doc.Saved Then Exit Sub
    Dim answer As VbMsgBoxResult
    answer = MsgBox("You have changed this wording draft." & vbCrLf & vbCrLf & _
                    "Save it to your Clause Library?" & vbCrLf & vbCrLf & _
                    "Yes  - save it, keeping the version it replaces" & vbCrLf & _
                    "No   - close and leave your library exactly as it is" & vbCrLf & _
                    "Cancel - go back to the draft", _
                    vbYesNoCancel + vbQuestion, "Clause Library draft")
    If answer = vbCancel Then Cancel = True: Exit Sub
    If answer = vbYes Then
        If Not CLSaveWordingFor(doc) Then Cancel = True: Exit Sub
    End If
    doc.Saved = True
End Sub
