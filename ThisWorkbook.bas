Option Explicit

Private mObjLogger As csLogger

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not mObjLogger Is Nothing Then
        mObjLogger.LogEventAction ("CLOSE")
        Set mObjLogger = Nothing
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    
    If Not mObjLogger Is Nothing Then
        mObjLogger.LogEventAction ("SAVE")
    End If

End Sub

Private Sub Workbook_Open()
    
    Set mObjLogger = New csLogger
    mObjLogger.LogEventAction ("OPEN")
    
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    If Not mObjLogger Is Nothing Then
        mObjLogger.LogSheetChangeEvent Sh, Target
    End If

End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not mObjLogger Is Nothing Then
        mObjLogger.LogSheetSelectionChangeEvent Sh, Target
    End If
End Sub