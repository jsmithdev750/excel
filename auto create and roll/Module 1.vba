Option Explicit
'=================================================
' Vanir File Manager - Fully working with fallback
'=================================================
Public Sub Vanir_File_Manager()

    Dim ws As Worksheet
    Dim fso As Object
    
    Dim userName As String
    Dim fizzPath As String, futurePath As String
    Dim yearFolder As String
    Dim todayDate As String
    Dim dateFormat As String
    Dim createdCount As Long
    
    ' File prefixes for first path
    Dim fizzOldCurve As String, fizzNewCurve As String
    
    ' File prefixes for second path
    Dim futureTradeList As String, futureOldCurve As String, futureNewCurve As String
    
    ' Variable to store source file once found
    Dim sourceFile As String
    
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
    todayDate = Format(CDate(ws.Range("D2").Value), dateFormat)
    
    '======================================
    ' Process Fizz path
    '======================================
    If fizzPath <> "" Then
        If Right(fizzPath, 1) <> "\" Then fizzPath = fizzPath & "\"
        fizzPath = "C:\Users\" & userName & "\" & fizzPath
        yearFolder = fizzPath & "Vanir JPN Fizz Curve Archive " & Year(Date) & "\"
        
        If Not fso.FolderExists(yearFolder) Then fso.CreateFolder yearFolder
        
        ' Process Fizz files using shared sourceFile
        sourceFile = ""
        If ProcessFileWithFallback(fso, yearFolder, fizzOldCurve, "", todayDate, ws.Range("D2").Value) Then createdCount = createdCount + 1
        If ProcessFileWithFallback(fso, yearFolder, fizzOldCurve, fizzNewCurve, todayDate, ws.Range("D2").Value) Then createdCount = createdCount + 1
        
        
    End If
    
    '======================================
    ' Process Future path
    '======================================
    If futurePath <> "" Then
        If Right(futurePath, 1) <> "\" Then futurePath = futurePath & "\"
        futurePath = "C:\Users\" & userName & "\" & futurePath
        yearFolder = futurePath & "Vanir JPN Curve Archive " & Year(Date) & "\"
        
        If Not fso.FolderExists(yearFolder) Then fso.CreateFolder yearFolder
        
        ' Process Future files using same shared sourceFile
        sourceFile = ""
        If ProcessFileWithFallback(fso, yearFolder, futureTradeList, "", todayDate, ws.Range("D2").Value) Then createdCount = createdCount + 1
        If ProcessFileWithFallback(fso, yearFolder, futureOldCurve, "", todayDate, ws.Range("D2").Value) Then createdCount = createdCount + 1
        If ProcessFileWithFallback(fso, yearFolder, futureOldCurve, futureNewCurve, todayDate, ws.Range("D2").Value) Then createdCount = createdCount + 1
    End If
    
    '--------------------------------------
    ' Result message
    '--------------------------------------
    If createdCount = 0 Then
        MsgBox "Today's files already exist in all folders.", vbInformation
    Else
        MsgBox createdCount & " file(s) created.", vbInformation
    End If

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
                                         ByVal todayValue As Variant) As Boolean

    Dim todayFile As String
    Dim latestFile As String
    Dim latestMonthFolder As String
    Dim prevYearFolder As String
    
    ' Build today's file name
    If suffix = "" Then
        todayFile = folderPath & baseName & "_" & todayDate & ".xlsx"
    Else
        todayFile = folderPath & baseName & "_" & todayDate & " " & suffix & ".xlsx"
    End If
    
    ' Skip if already exists
    If fso.FileExists(todayFile) Then Exit Function
    
    ' Step 1: Check current year folder
    latestFile = GetLatestFile(fso, folderPath, baseName, suffix)
    
    ' Step 2: Check month folders inside current year
    If latestFile = "" Then
        latestMonthFolder = GetLatestMonthFolder(fso, folderPath)
        If latestMonthFolder <> "" Then
            latestFile = GetLatestFile(fso, latestMonthFolder, baseName, suffix)
        End If
    End If
    
    ' Step 3: Fallback to previous year
    If latestFile = "" Then
    
        prevYearFolder = Replace(folderPath, CStr(Year(Date)), CStr(Year(Date) - 1))
        
        If fso.FolderExists(prevYearFolder) Then
        
            latestMonthFolder = GetLatestMonthFolder(fso, prevYearFolder)
            
            If latestMonthFolder <> "" Then
                latestFile = GetLatestFile(fso, latestMonthFolder, baseName, suffix)
            End If
            
        End If
        
    End If
    
    ' If still nothing ? stop
    If latestFile = "" Then Exit Function
    
    '------ Resetting the tradelist or curve --------
    ' Copy correct file
    fso.CopyFile latestFile, todayFile
    ' If this is NEW FORMAT file, clear data
    If InStr(1, todayFile, "NEW FORMAT", vbTextCompare) > 0 Then
        ClearRowsFrom2 todayFile
    End If
    
    ' If Futures Tradelist update data
    If InStr(1, todayFile, "Tradelist", vbTextCompare) > 0 Then
        ProcessFutureTradeList todayFile, todayValue
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
            
            '---------------------------------
            ' FIX: prevent old/new mix up
            '---------------------------------
            
            ' Old curve: ignore NEW FORMAT files
            If suffix = "" Then
                If InStr(1, file.Name, "NEW FORMAT", vbTextCompare) > 0 Then GoTo SkipFile
            End If
            
            ' New format: only accept NEW FORMAT files
            If suffix <> "" Then
                If InStr(1, file.Name, suffix, vbTextCompare) = 0 Then GoTo SkipFile
            End If
            
            '---------------------------------
            
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

    Dim wb As Workbook
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wb = Workbooks.Open(filePath)
    
    For Each ws In wb.Worksheets
        ws.Rows("2:" & ws.Rows.Count).ClearContents
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

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim c As Range
    Dim r As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim todayText As String
    Dim sectionNames As Variant
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    todayText = Format(CDate(todayValue), "d mmm yyyy")
    
    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Worksheets(1)
    
    '---------------------------
    ' Update Date
    '---------------------------
    Set c = ws.Cells.Find("Date:", LookAt:=xlPart, MatchCase:=False)
    
    If Not c Is Nothing Then
        c.Offset(0, 1).Value = todayText
    End If
    
    ' Sections we need
    sectionNames = Array("FUTURES", "OPTIONS", "OTC - VGM ONLY")
    
    '---------------------------
    ' Process each section
    '---------------------------
    For i = LBound(sectionNames) To UBound(sectionNames)
        
        Set c = ws.Cells.Find(sectionNames(i), LookAt:=xlWhole, MatchCase:=False)
        
        If Not c Is Nothing Then
            
            ' move to Product row
            If Trim(ws.Cells(c.Row + 1, c.Column).Value) = "Product" Then

                startRow = c.Row + 2
                endRow = startRow
                
                ' find end of data
                Do
                    If ws.Cells(endRow, c.Column).Value = "" Then Exit Do
                    If ws.Cells(endRow, c.Column).Value = "OTC" _
                    Or ws.Cells(endRow, c.Column).Value = "VGM" _
                    Or ws.Cells(endRow, c.Column).Value = "OTC - VGM ONLY" Then Exit Do
                    endRow = endRow + 1
                Loop
                
                '==========================
                ' Clear 6 columns safely
                '==========================
                Dim rr As Long, cc As Long
                Dim cell As Range
                Dim lastCol As Long
                
                lastCol = c.Column + 5
                
                rr = startRow
                Do While rr < endRow
                    cc = c.Column
                    Do While cc <= lastCol
                        Set cell = ws.Cells(rr, cc)
                        If cell.MergeCells Then
                            ' Clear the whole merged block
                            cell.MergeArea.ClearContents
                            ' Skip to column after merge area
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
    
    Application.ScreenUpdating = True

End Sub

