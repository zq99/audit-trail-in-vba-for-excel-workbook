Option Explicit
Option Compare Text

Private Type udtLogEntry
    Date            As String * 22
    NewCellValue    As String * 30
    OldCellValue    As String * 30
    CellRef         As String * 15
    UserName        As String * 10
    SheetName       As String * 20
    NewFormula      As String * 40
    OldFormula      As String * 40
    ChangeType      As String * 12
End Type

Private mudtEntry                           As udtLogEntry
Private Const CSTR_CELL_ADJUSTMENT_TYPE     As String = "Cell"
Private Const CSTR_LOG_FILENAME_SUFFIX      As String = "_log.txt"

Public Sub LogSheetChangeEvent(ByVal Sh As Object, ByVal Target As Range)

On Error GoTo ERR_HANDLER:
    Dim strText As String

    If Not ThisWorkbook.ReadOnly Then
        If (Target.Rows.Count = 1) And (Target.Columns.Count = 1) Then
            mudtEntry.SheetName = CStr(Sh.Name)
            mudtEntry.CellRef = CStr(Target.Address)
            mudtEntry.ChangeType = CSTR_CELL_ADJUSTMENT_TYPE
            mudtEntry.Date = CStr(Now())
            mudtEntry.NewCellValue = CStr(Target.Value)
            mudtEntry.UserName = Environ("username")
            mudtEntry.NewFormula = CStr(Target.Formula)
            strText = BuildLogString(mudtEntry.Date, mudtEntry.NewCellValue, _
                                     mudtEntry.OldCellValue, mudtEntry.CellRef, _
                                     mudtEntry.UserName, mudtEntry.SheetName, mudtEntry.OldFormula, _
                                     mudtEntry.NewFormula, mudtEntry.ChangeType)
            Call fnAddToFile(strText)
        End If
    End If
EXIT_HERE:
    Exit Sub
ERR_HANDLER:
    GoTo EXIT_HERE
End Sub

Public Sub LogSheetSelectionChangeEvent(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Not ThisWorkbook.ReadOnly Then
        If (Target.Rows.Count = 1) And (Target.Columns.Count = 1) Then
             mudtEntry.OldCellValue = CStr(Target.Value)
             mudtEntry.OldFormula = CStr(Target.Formula)
        End If
    End If
End Sub

Public Sub LogEventAction(ByVal strEvent As String)
    Dim udtEntry As udtLogEntry
    udtEntry.Date = Now()
    udtEntry.ChangeType = strEvent
    udtEntry.UserName = Environ("username")
    If Not fnAddToFile(udtEntry.Date & "," & udtEntry.UserName & "," & udtEntry.ChangeType) Then
        Debug.Print "Failed to log event"
    End If
End Sub

Private Function fnAddToFile(ByVal strText As String) As Boolean
    
On Error GoTo ERR_HANDLER:
    Dim intHandle   As Integer
    Dim strFileName As String
    
    fnAddToFile = False
    If ThisWorkbook.ReadOnly Then
        fnAddToFile = False
        GoTo EXIT_HERE
    End If
    
    intHandle = FreeFile
    strFileName = Mid(ThisWorkbook.Name, 1, InStr(1, ThisWorkbook.Name, ".") - 1)
    strFileName = strFileName & CSTR_LOG_FILENAME_SUFFIX
    strFileName = ThisWorkbook.Path & Chr(92) & strFileName
    
    If Not IsLogFilePresent(strFileName) Then
        Open strFileName For Append As #intHandle
            Dim udtHeader As udtLogEntry
            Dim strTitles As String
            
            udtHeader.SheetName = "Sheet Name"
            udtHeader.Date = "Date & Time"
            udtHeader.CellRef = "Cell Ref"
            udtHeader.SheetName = "Sheetname"
            udtHeader.UserName = "UserName"
            udtHeader.NewCellValue = "New Value"
            udtHeader.OldCellValue = "Old Value"
            udtHeader.NewFormula = "New Value Formula"
            udtHeader.OldFormula = "Old Value Formula"
            udtHeader.ChangeType = "Type"
            
            strTitles = BuildLogString(udtHeader.Date, udtHeader.NewCellValue, _
                                       udtHeader.OldCellValue, udtHeader.CellRef, _
                                       udtHeader.UserName, udtHeader.SheetName, _
                                       udtHeader.OldFormula, udtHeader.NewFormula, _
                                       udtHeader.ChangeType)
            Print #intHandle, strTitles
            Print #intHandle, strText
        Close #intHandle
    Else
        Open strFileName For Append As #intHandle
            Print #intHandle, strText
        Close #intHandle
    End If

    fnAddToFile = True
EXIT_HERE:
    Exit Function
ERR_HANDLER:
    fnAddToFile = False
    GoTo EXIT_HERE
End Function

Private Function BuildLogString(ByVal strDate As String, ByVal strNew As String, ByVal strOld As String, _
                                ByVal strRef As String, ByVal strName As String, ByVal strSheet As String, _
                                ByVal strOldFormula As String, ByVal strNewFormula As String, ByVal strChangeType As String) As String
    Dim strText As String
    On Error Resume Next
    strSheet = UCase(strSheet)
    BuildLogString = _
           strDate & "," & strName & "," & strChangeType & "," & strSheet & "," & strRef & ", " & strNew & "," & strOld & _
               "," & strNewFormula & "," & strOldFormula
End Function

Private Function IsLogFilePresent(ByVal strFile As String) As Boolean
On Error GoTo ERR_HANDLER:
    IsLogFilePresent = False
    If Trim(Dir(strFile)) <> "" Then
        IsLogFilePresent = True
    Else
        IsLogFilePresent = False
    End If
EXIT_HERE:
    Exit Function
ERR_HANDLER:
    IsLogFilePresent = False
    GoTo EXIT_HERE
End Function