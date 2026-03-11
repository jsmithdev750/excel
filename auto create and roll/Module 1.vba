Option Explicit
'=================================================
' Vanir File Manager - Fully working with fallback
'=================================================
Public Sub Vanir_File_Manager()

    Dim ws As Worksheet
    Dim fso As Object
    Dim workingDate As Date
    
    Dim userName As String
    Dim fizzPath As String, futurePath As String
    Dim yearFolder As String
    Dim todayDate As String
    Dim dateFormat As String
    
    ' File prefixes for first path
    Dim fizzOldCurve As String, fizzNewCurve As String
    
    ' File prefixes for second path
    Dim futureTradeList As String, futureOldCurve As String, futureNewCurve As String
    
    ' Track created/edited files
    Dim createdFiles As New Collection
    Dim editedFiles As New Collection
    
    
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '--------------------------------------
    ' Read config from sheet
    '--------------------------------------
    userName = Trim(ws.Range("D4").Value)
    dateFormat = ws.Range("A4").Value
    
    ' --- Fizz ---
    fizzPath = Trim(ws.Range("A15").Value)
    fizzOldCurve = Trim(ws.Range("B29").Value)  ' Base file
    fizzNewCurve = Trim(ws.Range("C29").Value)  ' New format
    
    ' --- Future ---
    futurePath = Trim(ws.Range("A15").Value)
    futureTradeList = Trim(ws.Range("B25").Value)  ' Base file
    futureOldCurve = Trim(ws.Range("C25").Value)   ' Second file
    futureNewCurve = Trim(ws.Range("D25").Value)   ' NEW FORMAT
    
    If userName = "" Or (fizzPath = "" And futurePath = "") Then
        MsgBox "Missing username or paths.", vbCritical
        Exit Sub
    End If
    
    '--------------------------------------
    ' Today's date
    '--------------------------------------
    workingDate = CDate(ws.Range("D2").Value)
    todayDate = Format(workingDate, dateFormat)
    
    '======================================
    ' Process Fizz path
    '======================================
    If fizzPath <> "" Then
        If Right(fizzPath, 1) <> "\" Then fizzPath = fizzPath & "\"
        fizzPath = "C:\Users\" & userName & "\" & fizzPath
        yearFolder = fizzPath & "Vanir JPN Fizz Curve Archive " & Year(workingDate) & "\"
        
        If Not fso.FolderExists(yearFolder) Then fso.CreateFolder yearFolder
        
        ' Process Fizz files
        ProcessFileWithFallback fso, yearFolder, fizzOldCurve, "", todayDate, workingDate, createdFiles, editedFiles
        ProcessFileWithFallback fso, yearFolder, fizzOldCurve, fizzNewCurve, todayDate, workingDate, createdFiles, editedFiles
    End If
    
    '======================================
    ' Process Future path
    '======================================
    If futurePath <> "" Then
        If Right(futurePath, 1) <> "\" Then futurePath = futurePath & "\"
        futurePath = "C:\Users\" & userName & "\" & futurePath
        yearFolder = futurePath & "Vanir JPN Curve Archive " & Year(workingDate) & "\"
        
        If Not fso.FolderExists(yearFolder) Then fso.CreateFolder yearFolder
        
        ' Process Future files
        ProcessFileWithFallback fso, yearFolder, futureTradeList, "", todayDate, workingDate, createdFiles, editedFiles
        ProcessFileWithFallback fso, yearFolder, futureOldCurve, "", todayDate, workingDate, createdFiles, editedFiles
        ProcessFileWithFallback fso, yearFolder, futureOldCurve, futureNewCurve, todayDate, workingDate, createdFiles, editedFiles
    End If
    
    '--------------------------------------
    ' Show summary (only file names)
    '--------------------------------------
    Dim msg As String
    Dim f As Variant
    
    msg = "Summary of today's run:" & vbCrLf & vbCrLf
    
    If createdFiles.Count > 0 Then
        msg = msg & "Created files:" & vbCrLf
        For Each f In createdFiles
            msg = msg & " - " & Dir(f) & vbCrLf
        Next f
    Else
        msg = msg & "No new files created." & vbCrLf
    End If
    
    If editedFiles.Count > 0 Then
        msg = msg & vbCrLf & "Edited workbooks:" & vbCrLf
        For Each f In editedFiles
            msg = msg & " - " & Dir(f) & vbCrLf
        Next f
    Else
        msg = msg & vbCrLf & "No workbooks edited."
    End If
    
    MsgBox msg, vbInformation, "Vanir Daily Summary"

ExitProc:
    Set fso = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume ExitProc

End Sub

'=================================================
' Process file with fallback
'=================================================
Private Function ProcessFileWithFallback(ByVal fso As Object, _
                                         ByVal folderPath As String, _
                                         ByVal baseName As String, _
                                         ByVal suffix As String, _
                                         ByVal todayDate As String, _
                                         ByVal todayValue As Variant, _
                                         ByRef createdFiles As Collection, _
                                         ByRef editedFiles As Collection) As Boolean

    Dim todayFile As String
    Dim latestFile As String
    Dim latestMonthFolder As String
    Dim prevYearFolder As String
    Dim didCloneToday As Boolean
    
    ' Build today's file name
    If suffix = "" Then
        todayFile = folderPath & baseName & "_" & todayDate & ".xlsx"
    Else
        todayFile = folderPath & baseName & "_" & todayDate & " " & suffix & ".xlsx"
    End If
    
    ' Step 1: Find the latest source file
    latestFile = GetLatestFile(fso, folderPath, baseName, suffix)
    
    If latestFile = "" Then
        latestMonthFolder = GetLatestMonthFolder(fso, folderPath)
        If latestMonthFolder <> "" Then latestFile = GetLatestFile(fso, latestMonthFolder, baseName, suffix)
    End If
    
    If latestFile = "" Then
        prevYearFolder = Replace(folderPath, CStr(Year(todayValue)), CStr(Year(todayValue) - 1))
        If fso.FolderExists(prevYearFolder) Then
            latestMonthFolder = GetLatestMonthFolder(fso, prevYearFolder)
            If latestMonthFolder <> "" Then latestFile = GetLatestFile(fso, latestMonthFolder, baseName, suffix)
        End If
    End If
    
    ' Exit if no source file found
    If latestFile = "" Then Exit Function
    
    ' ------------------------------------------------------
    ' CLONE today only if file doesn't exist
    ' ------------------------------------------------------
    didCloneToday = False
    If Not fso.FileExists(todayFile) Then
        fso.CopyFile latestFile, todayFile
        On Error Resume Next
        createdFiles.Add todayFile
        On Error GoTo 0
        didCloneToday = True
    End If
    
    ' ------------------------------------------------------
    ' Only edit if we cloned today
    ' ------------------------------------------------------
    If didCloneToday Then
        If InStr(1, todayFile, "NEW FORMAT", vbTextCompare) > 0 Then
            ClearRowsFrom2 todayFile
            editedFiles.Add todayFile
        End If
        
        If InStr(1, todayFile, "Tradelist", vbTextCompare) > 0 Then
            ProcessFutureTradeList todayFile, todayValue
            editedFiles.Add todayFile
        End If
        
        If InStr(1, todayFile, "Curve", vbTextCompare) > 0 _
        And InStr(1, todayFile, "PHYSICAL", vbTextCompare) > 0 _
        And InStr(1, todayFile, "NEW FORMAT", vbTextCompare) = 0 Then
            ResetPhysicalSheetsDaily todayFile
            editedFiles.Add todayFile
        End If
        
        ' Futures old curve (future logic)
        If InStr(1, todayFile, "Curve", vbTextCompare) > 0 _
        And InStr(1, todayFile, "PHYSICAL", vbTextCompare) = 0 _
        And InStr(1, todayFile, "NEW FORMAT", vbTextCompare) = 0 Then
            ' ResetFuturesCurve todayFile
        End If
    End If
    
    ProcessFileWithFallback = True
End Function

'=================================================
' Find latest file matching base name and optional suffix
'=================================================
Private Function GetLatestFile(ByVal fso As Object, _
                               ByVal folderPath As String, _
                               ByVal baseName As String, _
                               ByVal suffix As String) As String
    Dim folder As Object, file As Object
    Dim newestDate As Date, newestFile As String
    
    If Not fso.FolderExists(folderPath) Then Exit Function
    Set folder = fso.GetFolder(folderPath)
    newestDate = #1/1/1900#
    
    For Each file In folder.Files
        If InStr(1, file.Name, baseName, vbTextCompare) > 0 Then
            If suffix = "" Then
                If InStr(1, file.Name, "NEW FORMAT", vbTextCompare) > 0 Then GoTo SkipFile
            Else
                If InStr(1, file.Name, suffix, vbTextCompare) = 0 Then GoTo SkipFile
            End If
            
            If file.DateLastModified > newestDate Then
                newestDate = file.DateLastModified
                newestFile = file.Path
            End If
        End If
SkipFile:
    Next file
    
    GetLatestFile = newestFile
End Function

'=================================================
' Get latest month folder in mmyyyy format inside a year folder
'=================================================
Private Function GetLatestMonthFolder(ByVal fso As Object, ByVal yearFolder As String) As String
    Dim folder As Object, subFolder As Object
    Dim latestDate As Date, latestFolder As String
    
    If Not fso.FolderExists(yearFolder) Then Exit Function
    Set folder = fso.GetFolder(yearFolder)
    latestDate = #1/1/1900#
    
    For Each subFolder In folder.SubFolders
        If Len(subFolder.Name) = 6 And IsNumeric(subFolder.Name) Then
            Dim folderMonth As Integer, folderYear As Integer
            folderMonth = CInt(Left(subFolder.Name, 2))
            folderYear = CInt(Right(subFolder.Name, 4))
            
            If DateSerial(folderYear, folderMonth, 1) > latestDate Then
                latestDate = DateSerial(folderYear, folderMonth, 1)
                latestFolder = subFolder.Path & "\"
            End If
        End If
    Next subFolder
    
    GetLatestMonthFolder = latestFolder
End Function

'=================================================
' Clear rows from row 2 onwards
'=================================================
Private Sub ClearRowsFrom2(ByVal filePath As String)
    Dim wb As Workbook, ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wb = Workbooks.Open(filePath)
    
    For Each ws In wb.Worksheets
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then ws.Rows("2:" & lastRow).ClearContents
    Next ws
    
    wb.Save
    wb.Close False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'=================================================
' Update Futures Tradelist content
'=================================================
Private Sub ProcessFutureTradeList(ByVal filePath As String, ByVal todayValue As Variant)
    Dim wb As Workbook, ws As Worksheet, c As Range
    Dim startRow As Long, endRow As Long, todayText As String
    Dim sectionNames As Variant, i As Long
    On Error GoTo CleanExit
    Application.ScreenUpdating = False
    
    todayText = Format(CDate(todayValue), "d mmm yyyy")
    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Worksheets(1)
    
    Set c = ws.Cells.Find("Date:", LookAt:=xlPart, MatchCase:=False)
    If Not c Is Nothing Then c.Offset(0, 1).Value = todayText
    
    sectionNames = Array("FUTURES", "OPTIONS", "OTC - VGM ONLY")
    
    For i = LBound(sectionNames) To UBound(sectionNames)
        Set c = ws.Cells.Find(sectionNames(i), LookAt:=xlWhole, MatchCase:=False)
        If Not c Is Nothing Then
            If Trim(ws.Cells(c.Row + 1, c.Column).Value) = "Product" Then
                startRow = c.Row + 2
                endRow = startRow
                Do
                    If ws.Cells(endRow, c.Column).Value = "" Then Exit Do
                    If ws.Cells(endRow, c.Column).Value = "OTC" _
                    Or ws.Cells(endRow, c.Column).Value = "VGM" _
                    Or ws.Cells(endRow, c.Column).Value = "OTC - VGM ONLY" Then Exit Do
                    endRow = endRow + 1
                Loop
                
                Dim rr As Long, cc As Long, cell As Range, lastCol As Long
                lastCol = c.Column + 5
                rr = startRow
                Do While rr < endRow
                    cc = c.Column
                    Do While cc <= lastCol
                        Set cell = ws.Cells(rr, cc)
                        If cell.MergeCells Then
                            cell.MergeArea.ClearContents
                            cc = cc + cell.MergeArea.Columns.Count
                        Else
                            cell.ClearContents
                            cc = cc + 1
                        End If
                    Loop
                    rr = rr + 1
                Loop
            End If
        End If
    Next i
    
    wb.Save
    wb.Close False
    
CleanExit:
    Application.ScreenUpdating = True
End Sub

'=================================================
' Reset Fizz Curve Sheet Content (Physical Sheets)
' Returns collection of edited files
'=================================================
Private Sub ResetPhysicalSheetsDaily(ByVal filePath As String)

    Dim wb As Workbook, ws As Worksheet, cell As Range
    Dim lastCol As Long, newCol As Long, lastRow As Long, r As Long
    Dim todayDate As String, clrE2EFDA As Long
    
    On Error GoTo CleanExit
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'todayDate = Format(Date, "dd-mmm-yy")
    todayDate = Format(CDate(ThisWorkbook.Worksheets("Sheet1").Range("D2").Value), "dd-mmm-yy")
    clrE2EFDA = RGB(226, 239, 218)
    
    'Open workbook
    Set wb = Workbooks.Open(filePath)
    
    'Check workbook opened
    If wb Is Nothing Then GoTo CleanExit
    
    For Each ws In wb.Worksheets
        
        If InStr(1, ws.Name, "curve", vbTextCompare) = 0 Then
        
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            newCol = lastCol + 1
            
            ws.Columns(lastCol).Copy
            ws.Columns(newCol).PasteSpecial xlPasteAll
            
            ws.Cells(1, newCol).Value = todayDate
            
            lastRow = ws.Cells(ws.Rows.Count, newCol).End(xlUp).Row
            
            For r = 2 To lastRow
            
                Set cell = ws.Cells(r, newCol)
                
                If Not cell.HasFormula Then
                    If cell.Font.Color <> vbRed Then
                        If cell.Interior.Color = xlNone _
                        Or cell.Interior.Color = RGB(255, 255, 255) _
                        Or cell.Interior.Color = clrE2EFDA Then
                        
                            cell.ClearContents
                        
                        End If
                    End If
                End If
                
            Next r
            
        End If
        
    Next ws
    
    wb.Save
    wb.Close False

CleanExit:

    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
