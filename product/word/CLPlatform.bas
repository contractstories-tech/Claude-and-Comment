Attribute VB_Name = "CLPlatform"
' Clause Library - platform primitives: identity, time, text safety, files, hashing, XML.
' Nothing in this module knows what a clause is.
Option Explicit

Public Const CL_VERSION As String = "2.1"
Public Const CL_SCHEMA As String = "2"
Public Const CL_PRODUCT As String = "ClauseLibraryPersonal"
Public Const CL_ERR As Long = 2000          ' vbObjectError + CL_ERR = our own messages

' Windows refuses paths past 259 characters for ordinary file APIs. Rather than
' guess a root limit and hope, the budget is worked out from the longest path
' this product can ever generate, so the two numbers cannot drift apart:
'   \history\ (9) + id (36) + -r (2) + revision (10) + .rich (5)  = 62
'   + pending suffix  . (1) + 8 hex (8) + .pending (8)              = 17
Public Const CL_PATH_MAX As Long = 259
Public Const CL_PATH_RESERVE As Long = 79
Public Const CL_ROOT_MAX As Long = CL_PATH_MAX - CL_PATH_RESERVE      ' = 180

Private Type GUID
    a As Long
    b As Integer
    c As Integer
    d(0 To 7) As Byte
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef g As GUID) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef g As GUID, ByVal buffer As LongPtr, ByVal count As Long) As Long
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32" (ByRef provider As LongPtr, ByVal container As LongPtr, ByVal providerName As LongPtr, ByVal providerType As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32" (ByVal provider As LongPtr, ByVal algorithm As Long, ByVal key As LongPtr, ByVal flags As Long, ByRef hash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32" (ByVal hash As LongPtr, ByRef bytes As Any, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32" (ByVal hash As LongPtr, ByVal param As Long, ByRef bytes As Any, ByRef length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32" (ByVal hash As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32" (ByVal provider As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal existing As LongPtr, ByVal destination As LongPtr, ByVal flags As Long) As Long

Private Declare PtrSafe Sub SleepApi Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)

Private cachedFso As Object
Private stripper As Object

' ---------- errors ----------

' Raise a message written for the person using Word, not for a developer.
Public Sub CLFail(ByVal message As String)
    Err.Raise vbObjectError + CL_ERR, "ClauseLibrary", message
End Sub

' Turn any error into something a lawyer can act on. Word's own numbers leak
' otherwise: 70 appears as "Permission denied", 5941 as "The requested member
' of the collection does not exist."
Public Function CLExplain(ByVal number As Long, ByVal description As String) As String
    Select Case number
        Case 70
            CLExplain = "Your library is busy. Another Word window is saving to it right now. Wait a moment and try again."
        Case 5941, 4605
            CLExplain = "This document is not a Clause Library wording draft. Open the item in Library and choose Edit wording in Word first."
        Case 76, 53
            CLExplain = "Your library folder could not be reached. If it is on a removable or network drive, reconnect it, then use Locate a library."
        Case 75, 55
            CLExplain = "A library file is in use by another program. Close anything that has the library folder open and try again."
        Case 7, 14
            CLExplain = "Word ran out of memory for this operation. Close other documents and try again; very large selections are the usual cause."
        Case Else
            CLExplain = description
    End Select
    If Len(CLExplain) = 0 Then CLExplain = "An unexpected problem stopped this action (error " & number & ")."
End Function

' ---------- identity and time ----------

Public Function CLId() As String
    Dim g As GUID, b As String
    If CoCreateGuid(g) <> 0 Then CLFail "Windows could not create an identifier for this item."
    b = String$(39, vbNullChar)
    If StringFromGUID2(g, StrPtr(b), 39) = 0 Then CLFail "Windows could not create an identifier for this item."
    CLId = LCase$(Mid$(b, 2, 36))
End Function

Public Function CLIsId(ByVal id As String) As Boolean
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$"
    CLIsId = re.Test(id)
End Function

' One reading of the clock, and separators that cannot be rewritten by the
' machine's regional settings. Sorts correctly as text.
Public Function CLStamp(Optional ByVal when As Date = 0) As String
    Dim t As Date: t = IIf(when = 0, Now, when)
    CLStamp = Format$(Year(t), "0000") & "-" & Format$(Month(t), "00") & "-" & Format$(Day(t), "00") & _
              "T" & Format$(Hour(t), "00") & ":" & Format$(Minute(t), "00") & ":" & Format$(Second(t), "00")
End Function

' Reads back what CLStamp wrote, without depending on regional settings.
Public Function CLParseStamp(ByVal stamp As String) As Date
    On Error Resume Next
    If Len(stamp) < 19 Then Exit Function
    CLParseStamp = DateSerial(Val(Mid$(stamp, 1, 4)), Val(Mid$(stamp, 6, 2)), Val(Mid$(stamp, 9, 2))) _
                 + TimeSerial(Val(Mid$(stamp, 12, 2)), Val(Mid$(stamp, 15, 2)), Val(Mid$(stamp, 18, 2)))
    On Error GoTo 0
End Function

' "2026-09-19T14:02:33" -> "19 Sep 2026" for display.
Public Function CLNiceDate(ByVal stamp As String) As String
    If Len(stamp) < 10 Then CLNiceDate = "": Exit Function
    Dim y As Long, m As Long, d As Long
    y = Val(Mid$(stamp, 1, 4)): m = Val(Mid$(stamp, 6, 2)): d = Val(Mid$(stamp, 9, 2))
    If y = 0 Or m < 1 Or m > 12 Or d < 1 Then CLNiceDate = stamp: Exit Function
    CLNiceDate = d & " " & Mid$("JanFebMarAprMayJunJulAugSepOctNovDec", (m - 1) * 3 + 1, 3) & " " & y
End Function

' ---------- text safety ----------

' Word's Range.Text carries control characters that XML 1.0 simply cannot hold -
' footnote marks, inline picture anchors, page and column breaks, non-breaking
' and optional hyphens. Escaping does not help: &#1; is as illegal as Chr(1).
' Every string that reaches a library file passes through here first.
Public Function CLSafeText(ByVal value As String) As String
    If Len(value) = 0 Then CLSafeText = "": Exit Function
    value = Replace$(value, Chr$(7), vbTab)          ' table cell / row end
    value = Replace$(value, Chr$(11), vbLf)          ' manual line break
    value = Replace$(value, Chr$(12), vbLf)          ' page or section break
    value = Replace$(value, Chr$(14), vbLf)          ' column break
    value = Replace$(value, Chr$(30), "-")           ' non-breaking hyphen
    value = Replace$(value, Chr$(31), "")            ' optional (soft) hyphen
    value = Replace$(Replace$(value, vbCrLf, vbLf), vbCr, vbLf)
    ' Everything still below 0x20 other than tab and line feed is illegal in XML
    ' 1.0 and cannot be escaped into legality. Footnote marks, picture anchors
    ' and field characters all land here.
    If stripper Is Nothing Then
        Set stripper = CreateObject("VBScript.RegExp")
        stripper.Global = True
        stripper.Pattern = "[\x00-\x08\x0B-\x0C\x0E-\x1F\uFFFE\uFFFF]"
    End If
    CLSafeText = stripper.Replace(value, "")
End Function

' Trailing blank lines are noise in a clause; leading indentation is not.
Public Function CLTrimWording(ByVal value As String) As String
    Do While Len(value) > 0
        If Right$(value, 1) <> vbLf And Right$(value, 1) <> " " And Right$(value, 1) <> vbTab Then Exit Do
        value = Left$(value, Len(value) - 1)
    Loop
    CLTrimWording = value
End Function

' ---------- the writer lock ----------

' One writer at a time, across every Word process on this machine. Another
' window saving takes milliseconds, so waiting briefly is almost always
' invisible; failing instantly, as the first version did, was not.
Public Function CLTakeLock(ByVal root As String, ByRef fileNumber As Integer) As Boolean
    Dim attempt As Long
    For attempt = 1 To 12
        On Error Resume Next
        Err.Clear
        fileNumber = FreeFile
        Open root & "\writer.lock" For Binary Access Read Write Lock Read Write As #fileNumber
        If Err.number = 0 Then
            On Error GoTo 0
            CLTakeLock = True
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
        fileNumber = 0
        SleepApi 125
    Next
End Function

Public Sub CLRequireLock(ByVal root As String, ByRef fileNumber As Integer)
    If CLTakeLock(root, fileNumber) Then Exit Sub
    CLFail "Your library is busy. Another Word window has been saving to it for a few seconds." & vbCrLf & vbCrLf & _
           "Nothing was changed. Wait a moment and try again; if it keeps happening, close other Word windows."
End Sub

Public Sub CLDropLock(ByRef fileNumber As Integer)
    If fileNumber = 0 Then Exit Sub
    On Error Resume Next
    Close #fileNumber
    On Error GoTo 0
    fileNumber = 0
End Sub

' ---------- remembered preferences ----------

' Per-user, in this Windows profile. Never anything confidential - just which
' way round the person likes things.
Public Function CLSetting(ByVal name As String, ByVal fallback As String) As String
    CLSetting = GetSetting(CL_PRODUCT, "Options", name, fallback)
End Function

Public Sub CLSetSetting(ByVal name As String, ByVal value As String)
    SaveSetting CL_PRODUCT, "Options", name, value
End Sub

' ---------- keeping data where it was put ----------

' This product never opens a network connection. These two guards make sure a
' path it was handed cannot turn a local operation into a remote one: Word will
' happily open a web address, and explorer.exe will happily launch a browser.
' A drive letter or a \\server\share path is fine - a UNC path stays inside the
' network it belongs to. Anything with a scheme in it is refused.
Public Function CLIsLocalPath(ByVal path As String) As Boolean
    If Len(path) < 3 Then Exit Function
    If InStr(path, "://") > 0 Then Exit Function
    If Left$(path, 2) = "\\" Then CLIsLocalPath = True: Exit Function
    If Mid$(path, 2, 2) = ":\" And UCase$(Left$(path, 1)) >= "A" And UCase$(Left$(path, 1)) <= "Z" Then CLIsLocalPath = True
End Function

Public Sub CLRequireLocalPath(ByVal path As String, ByVal what As String)
    If CLIsLocalPath(path) Then Exit Sub
    CLFail "Clause Library only works with folders on this computer or your network. It will not use an internet address." _
        & vbCrLf & vbCrLf & what & ":" & vbCrLf & path
End Sub

' Names the file-sync service a folder belongs to, or "" if it is not in one.
' The tool itself sends nothing anywhere - but a folder inside OneDrive,
' Dropbox or similar is copied out by that service, which is the one way clause
' wording and private notes can leave this machine without anyone intending it.
Public Function CLSyncService(ByVal path As String) As String
    Dim probe As String, pair As Variant, parts As Variant, name As Variant
    probe = LCase$(path)
    For Each name In Array("OneDrive", "OneDriveCommercial", "OneDriveConsumer")
        If Len(Environ$(CStr(name))) > 0 Then
            If InStr(probe, LCase$(Environ$(CStr(name)))) = 1 Then CLSyncService = "OneDrive": Exit Function
        End If
    Next
    For Each pair In Array("\onedrive|OneDrive", "\dropbox|Dropbox", "\box|Box", _
                           "\google drive|Google Drive", "\googledrive|Google Drive", _
                           "\my drive|Google Drive", "\icloud|iCloud", _
                           "\nextcloud|Nextcloud", "\owncloud|ownCloud", "\egnyte|Egnyte", _
                           "\creative cloud files|Creative Cloud", "\pcloud|pCloud", "\sharepoint|SharePoint")
        parts = Split(CStr(pair), "|")
        If InStr(probe, CStr(parts(0))) > 0 Then CLSyncService = CStr(parts(1)): Exit Function
    Next
End Function

' Returns False if the person decides against the folder. Asked once per
' service; the answer is remembered so it does not become background noise.
Public Function CLConfirmSyncedFolder(ByVal path As String, ByVal what As String) As Boolean
    CLConfirmSyncedFolder = True
    Dim service As String: service = CLSyncService(path)
    If Len(service) = 0 Then Exit Function
    If CLSetting("SyncFolderAccepted", "") = service Then Exit Function
    CLConfirmSyncedFolder = (MsgBox( _
        "This folder is inside " & service & ":" & vbCrLf & vbCrLf & "    " & path & vbCrLf & vbCrLf & _
        what & " will therefore be copied to " & service & " by " & service & " itself, including your private notes." & vbCrLf & vbCrLf & _
        "Clause Library never sends anything anywhere. This is the one way your wording can leave this computer " & _
        "without you meaning it to, so it is worth being sure your organisation's data rules allow it." & vbCrLf & vbCrLf & _
        "Use this folder anyway?", vbYesNo + vbExclamation + vbDefaultButton2, "This folder syncs to " & service) = vbYes)
    If CLConfirmSyncedFolder Then CLSetSetting "SyncFolderAccepted", service
End Function

Public Function CLHtml(ByVal value As String) As String
    value = Replace$(value, "&", "&amp;")
    value = Replace$(value, "<", "&lt;"): value = Replace$(value, ">", "&gt;")
    value = Replace$(value, Chr$(34), "&quot;"): value = Replace$(value, "'", "&#39;")
    CLHtml = value
End Function

Public Function CLCsv(ByVal value As String) As String
    CLCsv = Chr$(34) & Replace$(Replace$(CLSafeText(value), Chr$(34), Chr$(34) & Chr$(34)), vbLf, " ") & Chr$(34)
End Function

' A file name Windows will accept, derived from something a person typed.
Public Function CLSafeFileName(ByVal value As String, ByVal maximum As Long) As String
    Dim bad As Variant, b As Variant
    value = CLSafeText(value)
    For Each b In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|", vbTab, vbLf)
        value = Replace$(value, CStr(b), " ")
    Next
    Do While InStr(value, "  ") > 0: value = Replace$(value, "  ", " "): Loop
    value = Trim$(value)
    Do While Len(value) > 0 And Right$(value, 1) = "."
        value = Left$(value, Len(value) - 1)
    Loop
    If Len(value) > maximum Then value = RTrim$(Left$(value, maximum))
    If Len(value) = 0 Then value = "Untitled"
    CLSafeFileName = value
End Function

' ---------- files ----------

Public Function CLFso() As Object
    If cachedFso Is Nothing Then Set cachedFso = CreateObject("Scripting.FileSystemObject")
    Set CLFso = cachedFso
End Function

Public Function CLExists(ByVal path As String) As Boolean
    CLExists = CLFso().FileExists(path)
End Function

Public Function CLFolderExists(ByVal path As String) As Boolean
    CLFolderExists = CLFso().FolderExists(path)
End Function

Public Sub CLFolder(ByVal path As String)
    If Not CLFso().FolderExists(path) Then CLFso().CreateFolder path
End Sub

Public Function CLRead(ByVal path As String) As String
    Dim s As Object: Set s = CreateObject("ADODB.Stream")
    s.Type = 2: s.Charset = "utf-8": s.Open
    s.LoadFromFile path: CLRead = s.ReadText: s.Close
End Function

Public Sub CLWrite(ByVal path As String, ByVal value As String)
    If Len(path) > CL_PATH_MAX Then
        CLFail "This file's path is longer than Windows allows (" & Len(path) & " characters)." & vbCrLf & vbCrLf & _
               "Move your library closer to the top of a drive and try again. Nothing was written."
    End If
    Dim s As Object: Set s = CreateObject("ADODB.Stream")
    s.Type = 2: s.Charset = "utf-8": s.Open: s.WriteText value: s.SaveToFile path, 2: s.Close
End Sub

' Write somewhere else, then swap it in as one filesystem operation. A crash
' at any point leaves either the old file or the new one, never half of either.
Public Sub CLAtomicWrite(ByVal path As String, ByVal value As String)
    Dim pending As String: pending = path & "." & Left$(CLId(), 8) & ".pending"
    On Error GoTo Failed
    CLWrite pending, value
    If CLRead(pending) <> value Then CLFail "This save could not be verified on disk. Your previous saved version was not replaced."
    If MoveFileExW(StrPtr(pending), StrPtr(path), 9) = 0 Then
        CLFail "Windows would not finish saving this file. Check that the library folder is available and not read-only. Your previous saved version was preserved."
    End If
    Exit Sub
Failed:
    Dim message As String: message = CLExplain(Err.number, Err.Description)
    On Error Resume Next
    If CLExists(pending) Then CLFso().DeleteFile pending
    On Error GoTo 0
    CLFail message
End Sub

' Left behind only if Windows refused a swap. Harmless, but they should not pile up.
Public Function CLSweepPending(ByVal folder As String) As Long
    Dim f As Object, removed As Long
    If Not CLFolderExists(folder) Then Exit Function
    On Error Resume Next
    For Each f In CLFso().GetFolder(folder).Files
        If LCase$(Right$(f.Name, 8)) = ".pending" Then
            If DateDiff("n", f.DateLastModified, Now) > 60 Then
                CLFso().DeleteFile f.path: removed = removed + 1
            End If
        End If
    Next
    On Error GoTo 0
    CLSweepPending = removed
End Function

Public Sub CLCopyFile(ByVal source As String, ByVal destination As String)
    CLFso().CopyFile source, destination, True
End Sub

' ---------- hashing ----------

Private Function CLHashBegin(ByRef provider As LongPtr, ByRef hash As LongPtr) As Boolean
    If CryptAcquireContextW(provider, 0, 0, 24, &HF0000000) = 0 Then Exit Function
    If CryptCreateHash(provider, &H800C&, 0, 0, hash) = 0 Then Exit Function
    CLHashBegin = True
End Function

Private Function CLHashEnd(ByVal hash As LongPtr) As String
    Dim digest(0 To 31) As Byte, size As Long, i As Long, out As String
    size = 32
    If CryptGetHashParam(hash, 2, digest(0), size, 0) = 0 Then Exit Function
    For i = 0 To 31: out = out & LCase$(Right$("0" & Hex$(digest(i)), 2)): Next
    CLHashEnd = out
End Function

Public Function CLSha(ByVal value As String) As String
    Dim p As LongPtr, h As LongPtr, b() As Byte, s As Object
    On Error GoTo Failed
    If Not CLHashBegin(p, h) Then CLFail "Windows could not start a checksum."
    If Len(value) > 0 Then
        Set s = CreateObject("ADODB.Stream")
        s.Type = 2: s.Charset = "utf-8": s.Open: s.WriteText value
        s.Position = 0: s.Type = 1: s.Position = 3: b = s.Read: s.Close   ' skip the UTF-8 byte order mark
        If CryptHashData(h, b(0), UBound(b) + 1, 0) = 0 Then CLFail "Windows could not compute a checksum."
    End If
    CLSha = CLHashEnd(h)
    CryptDestroyHash h: CryptReleaseContext p, 0
    Exit Function
Failed:
    Dim message As String: message = Err.Description
    On Error Resume Next
    If h <> 0 Then CryptDestroyHash h
    If p <> 0 Then CryptReleaseContext p, 0
    On Error GoTo 0
    CLFail message
End Function

' Hashes the bytes actually on disk, which is what integrity should mean.
Public Function CLFileSha(ByVal path As String) As String
    Dim p As LongPtr, h As LongPtr, fileNumber As Integer, opened As Boolean
    Dim remaining As Long, chunk As Long, bytes() As Byte
    On Error GoTo Failed
    If Not CLHashBegin(p, h) Then CLFail "Windows could not start a checksum."
    fileNumber = FreeFile
    Open path For Binary Access Read Shared As #fileNumber: opened = True
    remaining = LOF(fileNumber)
    Do While remaining > 0
        chunk = remaining: If chunk > 65536 Then chunk = 65536
        ReDim bytes(0 To chunk - 1): Get #fileNumber, , bytes
        If CryptHashData(h, bytes(0), chunk, 0) = 0 Then CLFail "Windows could not compute a checksum."
        remaining = remaining - chunk
    Loop
    Close #fileNumber: opened = False
    CLFileSha = CLHashEnd(h)
    CryptDestroyHash h: CryptReleaseContext p, 0
    Exit Function
Failed:
    Dim message As String: message = CLExplain(Err.number, Err.Description)
    On Error Resume Next
    If opened Then Close #fileNumber
    If h <> 0 Then CryptDestroyHash h
    If p <> 0 Then CryptReleaseContext p, 0
    On Error GoTo 0
    CLFail message
End Function

' ---------- XML ----------

Public Function CLDom(Optional ByVal xml As String = "") As Object
    Dim d As Object: Set d = CreateObject("MSXML2.DOMDocument.6.0")
    d.async = False: d.preserveWhiteSpace = True: d.validateOnParse = False: d.resolveExternals = False
    d.setProperty "ProhibitDTD", True
    If Len(xml) > 0 Then
        If Not d.LoadXML(xml) Then CLFail "This library file could not be read as valid XML. It has not been changed."
    End If
    Set CLDom = d
End Function
