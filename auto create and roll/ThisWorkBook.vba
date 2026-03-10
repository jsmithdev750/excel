Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ' --- Fill username into T6 ---
    Dim userName As String
    userName = Environ("Username")
    ThisWorkbook.Sheets("Sheet1").Range("D4").Value = userName

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description
    Resume CleanExit
End Sub

' Helper function
Private Function SheetExists(sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function



