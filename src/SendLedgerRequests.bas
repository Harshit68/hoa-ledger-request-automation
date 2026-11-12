Option Explicit

Sub SendEmails()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim toRecipients As String
    Dim ccRecipient As String
    Dim cell As Range

    ' === SETTINGS ===
    Const LOG_COL As String = "U"                 ' Where to log results
    Const INCLUDE_SUBFOLDERS As Boolean = False   ' True to search subfolders as well
    Const suffix As String = "( - Emailed 07.07 Detailed Ledger)" ' <--- ADDED: rename suffix

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change if needed

    ' Find the last row with data in column P (checkbox column)
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row

    ' Create Outlook Application object
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If OutlookApp Is Nothing Then
        MsgBox "Could not start Outlook. Please open Outlook and try again.", vbExclamation
        Exit Sub
    End If

    ' Resolve search folder: ALWAYS from A1
    Dim baseFolder As String
    baseFolder = Trim$(CStr(ws.Range("A1").Value))
    If Len(baseFolder) = 0 Then
        MsgBox "Folder path is blank in A1. Please enter a valid folder path.", vbExclamation
        Exit Sub
    End If
    ' Strip accidental quotes if pasted
    If InStr(baseFolder, """") > 0 Then baseFolder = Replace(baseFolder, """", "")
    baseFolder = Trim$(baseFolder)

    ' Validate folder exists
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseFolder) Then
        MsgBox "Folder not found: " & baseFolder, vbExclamation
        Exit Sub
    End If

    ' Loop through each row
    For i = 2 To lastRow ' Assuming row 1 has headers
        ' Check if the checkbox in column P is checked/true
        If IsChecked(ws.Cells(i, "P").Value) Then
            ' Collect email addresses from K and L for "To"
            toRecipients = ""
            For Each cell In ws.Range("K" & i & ":L" & i)
                If Trim$(CStr(cell.Value)) <> "" Then
                    If toRecipients = "" Then
                        toRecipients = Trim$(CStr(cell.Value))
                    Else
                        toRecipients = toRecipients & ";" & Trim$(CStr(cell.Value))
                    End If
                End If
            Next cell

            ' Get CC recipient from column M
            ccRecipient = Trim$(CStr(ws.Cells(i, "M").Value))

            ' Ensure there is at least one recipient
            If toRecipients <> "" Then
                ' Create a new email
                Set OutlookMail = OutlookApp.CreateItem(0)

                With OutlookMail
                    .To = toRecipients
                    If ccRecipient <> "" Then .CC = ccRecipient
                    .Subject = CStr(ws.Cells(i, "O").Value)
                    .Body = CStr(ws.Cells(i, "N").Value)

                    ' === Auto-attach the single latest PDF that matches ALL keywords in Column A ===
                    Dim kwRaw As String
                    Dim keywordsArr() As String
                    Dim bestPath As String
                    Dim attached As Boolean

                    kwRaw = CStr(ws.Cells(i, "A").Value)
                    keywordsArr = SplitKeywords(kwRaw)

                    If UBoundSafe(keywordsArr) >= 0 Then
                        bestPath = FindSingleLatestPdfWithAllKeywords(baseFolder, keywordsArr, INCLUDE_SUBFOLDERS)
                        If Len(bestPath) > 0 Then
                            On Error Resume Next
                            .Attachments.Add bestPath
                            attached = (Err.Number = 0)
                            Err.Clear
                            On Error GoTo 0
                        End If
                    End If

                    ' Log results to U
                    If attached Then
                        ws.Cells(i, LOG_COL).Value = "Attached: " & bestPath
                    Else
                        If Trim$(kwRaw) = "" Then
                            ws.Cells(i, LOG_COL).Value = "No keywords in Col A"
                        Else
                            ws.Cells(i, LOG_COL).Value = "No matching PDFs"
                        End If
                    End If
                    ' === end auto-attach ===

                    ' === Send or Display ===
                    ' NOTE: If you want the rename to happen only after *sending*,
                    ' use .Send (uncomment) and comment out .Display.
                    '.Send
                    .Display
                End With

                ' === Timestamp in Col T ===
                ws.Cells(i, "T").Value = Now
                ws.Cells(i, "T").NumberFormat = "mm/dd/yyyy hh:mm:ss"

                ' === NEW: Rename the attached file after send/display ===
                If Len(bestPath) > 0 Then
                    Dim newPath As String
                    newPath = RenameFileAppendSuffix(bestPath, suffix)
                    If Len(newPath) > 0 Then
                        ' Also reflect the rename in the log cell
                        ws.Cells(i, LOG_COL).Value = "Attached & Renamed: " & newPath
                    Else
                        ' If failed, keep prior log but add a note
                        ws.Cells(i, LOG_COL).Value = ws.Cells(i, LOG_COL).Value & " | Rename failed"
                    End If
                End If

                ' Clean up
                Set OutlookMail = Nothing
            Else
                ws.Cells(i, LOG_COL).Value = "Skipped: no To"
                ' Optional: clear timestamp for skipped rows
                ' ws.Cells(i, "T").ClearContents
            End If
        End If
    Next i

    ' Cleanup
    Set OutlookApp = Nothing
    MsgBox "Emails created successfully!", vbInformation
End Sub

' ---------- Helpers ----------

Private Function IsChecked(v As Variant) As Boolean
    On Error Resume Next
    Select Case True
        Case VarType(v) = vbBoolean: IsChecked = v
        Case IsNumeric(v): IsChecked = (v <> 0)
        Case TypeName(v) = "String"
            Dim s As String: s = UCase$(Trim$(v))
            IsChecked = (s = "TRUE" Or s = "YES" Or s = "Y" Or s = "X")
        Case Else
            IsChecked = False
    End Select
End Function

Private Function UBoundSafe(arr() As String) As Long
    On Error GoTo EmptyArr
    UBoundSafe = UBound(arr)
    Exit Function
EmptyArr:
    UBoundSafe = -1
End Function

Private Function SplitKeywords(ByVal s As String) As String()
    Dim t As String
    Dim arr As Variant, item As Variant, cleaned() As String
    Dim i As Long

    ' Normalize separators to semicolons
    t = Trim$(s)
    t = Replace(t, vbCrLf, ";")
    t = Replace(t, vbLf, ";")
    t = Replace(t, vbCr, ";")
    t = Replace(t, ",", ";")
    Do While InStr(t, ";;") > 0: t = Replace(t, ";;", ";"): Loop
    If Left$(t, 1) = ";" Then t = Mid$(t, 2)
    If Right$(t, 1) = ";" Then t = Left$(t, Len(t) - 1)

    If Len(t) = 0 Then
        ReDim cleaned(-1 To -1) ' empty array
        SplitKeywords = cleaned
        Exit Function
    End If
    arr = Split(t, ";")

    ' Deduplicate & trim (case-insensitive)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    For Each item In arr
        Dim k As String: k = LCase$(Trim$(CStr(item)))
        If Len(k) > 0 Then
            If Not dict.Exists(k) Then dict.Add k, k
        End If
    Next item

    If dict.Count = 0 Then
        ReDim cleaned(-1 To -1)
    Else
        ReDim cleaned(0 To dict.Count - 1)
        For i = 0 To dict.Count - 1
            cleaned(i) = CStr(dict.Items()(i))
        Next i
    End If

    SplitKeywords = cleaned
End Function

Private Function FindSingleLatestPdfWithAllKeywords( _
    ByVal baseFolder As String, _
    ByRef keywords() As String, _
    ByVal includeSubfolders As Boolean) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseFolder) Then
        FindSingleLatestPdfWithAllKeywords = ""
        Exit Function
    End If

    Dim bestPath As String
    Dim bestTime As Date: bestTime = CDate(0)

    Dim root As Object: Set root = fso.GetFolder(baseFolder)
    ProcessFolderForAllKws root, keywords, includeSubfolders, bestPath, bestTime
    FindSingleLatestPdfWithAllKeywords = bestPath
End Function

Private Sub ProcessFolderForAllKws( _
    ByVal fldr As Object, _
    ByRef keywords() As String, _
    ByVal includeSubfolders As Boolean, _
    ByRef bestPath As String, _
    ByRef bestTime As Date)

    Dim f As Object
    For Each f In fldr.Files
        ' Only PDFs
        If Len(f.Name) >= 4 Then
            If LCase$(Right$(f.Name, 4)) = ".pdf" Then
                Dim fnameLower As String
                fnameLower = LCase$(f.Name)
                If ContainsAllKeywords(fnameLower, keywords) Then
                    If CDate(f.DateLastModified) > bestTime Then
                        bestTime = CDate(f.DateLastModified)
                        bestPath = f.Path
                    End If
                End If
            End If
        End If
    Next f

    If includeSubfolders Then
        Dim subf As Object
        For Each subf In fldr.SubFolders
            ProcessFolderForAllKws subf, keywords, includeSubfolders, bestPath, bestTime
        Next subf
    End If
End Sub

Private Function ContainsAllKeywords(ByVal fileNameLower As String, ByRef keywords() As String) As Boolean
    Dim i As Long
    For i = LBound(keywords) To UBound(keywords)
        If InStr(1, fileNameLower, LCase$(keywords(i)), vbTextCompare) = 0 Then
            ContainsAllKeywords = False
            Exit Function
        End If
    Next i
    ContainsAllKeywords = True
End Function

' ===== NEW HELPERS FOR RENAMING =====

Private Function AppendSuffixBeforeExtension(ByVal fullPath As String, ByVal suffix As String) As String
    ' Returns a path where the filename is appended with suffix before its extension.
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fullPath) Then
        AppendSuffixBeforeExtension = ""
        Exit Function
    End If

    Dim folderPath As String, baseName As String, ext As String
    folderPath = fso.GetFile(fullPath).ParentFolder.Path
    baseName = fso.GetBaseName(fullPath)
    ext = fso.GetExtensionName(fullPath)

    ' If already has the suffix, return empty to signal "no change"
    If Right$(baseName, Len(suffix)) = suffix Then
        AppendSuffixBeforeExtension = "" ' no change necessary
        Exit Function
    End If

    ' Sanitize suffix for invalid filename characters
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In badChars
        suffix = Replace(suffix, CStr(ch), " ")
    Next ch
    suffix = Trim$(suffix)

    Dim newName As String
    If ext <> "" Then
        newName = baseName & suffix & "." & ext
    Else
        newName = baseName & suffix
    End If

    AppendSuffixBeforeExtension = fso.BuildPath(folderPath, newName)
End Function

Private Function RenameFileAppendSuffix(ByVal oldPath As String, ByVal suffix As String) As String
    On Error GoTo Fail
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(oldPath) Then
        RenameFileAppendSuffix = ""
        Exit Function
    End If

    Dim candidate As String
    candidate = AppendSuffixBeforeExtension(oldPath, suffix)

    ' If empty, either file missing or suffix already present
    If Len(candidate) = 0 Then
        RenameFileAppendSuffix = oldPath ' Treat as no-op
        Exit Function
    End If

    ' Handle collisions by adding (1), (2), ...
    If fso.FileExists(candidate) Then
        Dim n As Long
        n = 1
        Dim base As String, ext As String, folderPath As String
        Dim fileObj As Object: Set fileObj = fso.GetFile(oldPath)
        folderPath = fileObj.ParentFolder.Path
        base = fso.GetBaseName(candidate)
        ext = fso.GetExtensionName(candidate)

        Do While fso.FileExists(candidate)
            If ext <> "" Then
                candidate = fso.BuildPath(folderPath, base & " (" & n & ")." & ext)
            Else
                candidate = fso.BuildPath(folderPath, base & " (" & n & ")")
            End If
            n = n + 1
        Loop
    End If

    fso.MoveFile oldPath, candidate
    RenameFileAppendSuffix = candidate
    Exit Function

Fail:
    RenameFileAppendSuffix = ""
End Function




==================== Sheet3 doc
Attribute VB_Name = "Sheet3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
