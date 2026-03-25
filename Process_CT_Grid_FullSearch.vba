Option Explicit

Public Sub Process_CT_Grid_FullSearch()

    '-----------------------------------
    ' Declare variables
    '-----------------------------------
    Dim ws As Worksheet
    
    Dim lastRow As Long
    Dim lastCol As Long
    
    Dim contractCol As Long
    Dim headerRow As Long
    
    Dim i As Long
    Dim r As Long, c As Long
    
    Dim cellValue As String
    Dim parts As Variant
    
    Dim regionName As String
    Dim contractName As String
    Dim valueNum As Variant
    
    Dim regionCol As Long
    Dim contractRow As Long
    
    '-----------------------------------
    ' Setup worksheet
    '-----------------------------------
    Set ws = ThisWorkbook.Sheets("CT GRID Last value")
    
    '-----------------------------------
    ' Detect boundaries
    '-----------------------------------
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    '-----------------------------------
    ' STEP 1: Find "Contract" column (ANYWHERE)
    '-----------------------------------
    contractCol = 0
    
    For r = 1 To lastRow
        For c = 1 To lastCol
            If CleanText(ws.Cells(r, c).Value) = "contract" Then
                contractCol = c
                headerRow = r
                Exit For
            End If
        Next c
        If contractCol <> 0 Then Exit For
    Next r
    
    If contractCol = 0 Then
        MsgBox "Contract column not found.", vbCritical
        Exit Sub
    End If
    
    '-----------------------------------
    ' STEP 1.5: Find last row of Contract column
    '-----------------------------------
    Dim contractLastRow As Long
    contractLastRow = ws.Cells(ws.Rows.Count, contractCol).End(xlUp).Row
    
    
    '-----------------------------------
    ' STEP 1.6: Clear CT columns (ONLY after contractCol)
    '-----------------------------------
    Dim ctFound As Boolean
    Dim ctRow As Long
    
    For c = contractCol + 1 To lastCol
        
        ctFound = False
        
        ' Search down this column for "CT"
        For r = 1 To contractLastRow
            
            If Not IsError(ws.Cells(r, c).Value) Then
                
                If InStr(1, CleanText(ws.Cells(r, c).Value), "ct") > 0 Then
                    
                    ctRow = r
                    ctFound = True
                    Exit For
                    
                End If
                
            End If
            
        Next r
        
        ' If CT found in this column ? clear below it
        If ctFound Then
            
            If ctRow < contractLastRow Then
                
                ws.Range(ws.Cells(ctRow + 1, c), _
                         ws.Cells(contractLastRow, c)).ClearContents
                         
            End If
            
        End If
        
    Next c
    
    '-----------------------------------
    ' STEP 2: Loop input data (Column A)
    '-----------------------------------
    For i = 2 To lastRow
        
        cellValue = ws.Cells(i, 1).Value
        cellValue = CleanText(cellValue)
        
        If cellValue = "" Then GoTo NextRow
        
        parts = Split(cellValue, " ")
        If UBound(parts) < 2 Then GoTo NextRow
        
        regionName = parts(0)
        contractName = parts(1)
        valueNum = parts(2)
        
        '-----------------------------------
        ' STEP 3: Find REGION column (search whole sheet but clean)
        '-----------------------------------
        regionCol = 0
        
        For r = 1 To lastRow
            For c = 1 To lastCol
                If CleanText(ws.Cells(r, c).Value) = regionName Then
                    regionCol = c
                    Exit For
                End If
            Next c
            If regionCol <> 0 Then Exit For
        Next r
        
        If regionCol = 0 Then
            Debug.Print "Region NOT FOUND: " & regionName
            GoTo NextRow
        End If
        
        '-----------------------------------
        ' STEP 4: Find CONTRACT row (FIXED - ONLY search contract column)
        '-----------------------------------
        contractRow = 0
        
        For r = headerRow + 1 To lastRow
            If CleanText(ws.Cells(r, contractCol).Value) = contractName Then
                contractRow = r
                Exit For
            End If
        Next r
        
        If contractRow = 0 Then
            Debug.Print "Contract NOT FOUND: " & contractName
            GoTo NextRow
        End If
        
        '-----------------------------------
        ' STEP 5: Write value
        '-----------------------------------
        ws.Cells(contractRow, regionCol + 1).Value = valueNum
        
NextRow:
    Next i
    
    MsgBox "Processing completed.", vbInformation

End Sub

'-----------------------------------
' CLEAN FUNCTION (VERY IMPORTANT)
'-----------------------------------
Private Function CleanText(ByVal txt As String) As String
    
    If txt = "" Then
        CleanText = ""
        Exit Function
    End If
    
    ' Remove non-breaking spaces
    txt = replace(txt, Chr(160), " ")
    
    ' Remove non-printable characters
    txt = Application.WorksheetFunction.Clean(txt)
    
    ' Trim spaces
    txt = Trim(txt)
    
    ' Normalize multiple spaces
    Do While InStr(txt, "  ") > 0
        txt = replace(txt, "  ", " ")
    Loop
    
    CleanText = LCase(txt)

End Function
