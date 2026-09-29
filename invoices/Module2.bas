Attribute VB_Name = "Module2"
Sub ListSheetNamesFromWorkbook()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outputWs As Worksheet
    Dim settingsWS As Worksheet
    Dim basePath As String
    Dim sourcePath As String
    Dim i As Long
    Dim r As Long
    Dim lastRow As Long
    
    Call UpdateUsername
    'Settings sheet
    Set settingsWS = ThisWorkbook.Worksheets("Settings")
    
    'Build username path
    basePath = "C:\Users\" & settingsWS.Range("A5").value
    
    'Output sheet
    Set outputWs = ThisWorkbook.Worksheets("Sheet name and c2")
    
    'Clear old list
    outputWs.Range("A:A").ClearContents
    outputWs.Range("B:B").ClearContents
    
    i = 1
    
    'Loop through file paths in A3:A4
    lastRow = settingsWS.Cells(settingsWS.Rows.Count, "A").End(xlUp).Row
    
    For r = 8 To 9
        
        sourcePath = basePath & settingsWS.Cells(r, "A").value
        
        'Check file exists
        If Dir(sourcePath) <> "" Then
        
            Set wb = Workbooks.Open(sourcePath, ReadOnly:=True)
            
            'Add workbook name as header
            outputWs.Cells(i, 1).value = Left(wb.Name, InStrRev(wb.Name, ".") - 1)
            i = i + 1
            
            'List sheet names with filter
            For Each ws In wb.Worksheets
            
                If IsTargetSheet(ws.Name) Then
                    
                    outputWs.Cells(i, 1).value = ws.Name
                    outputWs.Cells(i, 2).value = GetC2Value(ws)
                    
                    i = i + 1
                    
                End If
            
            Next ws
            
            wb.Close SaveChanges:=False
            
            'Blank line between files
            i = i + 1
            
        Else
            outputWs.Cells(i, 1).value = "File not found: " & sourcePath
            i = i + 1
        End If
        
    Next r

End Sub
Function IsTargetSheet(sheetName As String) As Boolean

    Dim months As Variant
    Dim m As Variant
    Dim excludeList As Variant
    Dim e As Variant
    
    sheetName = LCase(sheetName)
    
    'Exclude specific keywords
    excludeList = Array( _
        "change jpysgd rate here", _
        "change invoice date here", _
        "template - carbon", _
        "sample invoice -kepco", _
        "enetrade cancelled trades lump", _
        "combined", _
        "buy", _
        "sell", _
        "sheet" _
    )
    
    For Each e In excludeList
        If InStr(sheetName, e) > 0 Then
            IsTargetSheet = False
            Exit Function
        End If
    Next e
    
    
    'Exclude month/date sheets (example: June 2026v2)
    months = Array("jan", "feb", "mar", "apr", "may", "jun", _
                   "jul", "aug", "sep", "oct", "nov", "dec", _
                   "january", "february", "march", "april", _
                   "june", "july", "august", "september", _
                   "october", "november", "december")
    
    For Each m In months
        If InStr(sheetName, m) > 0 Then
            If sheetName Like "*20##*" Then
                IsTargetSheet = False
                Exit Function
            End If
        End If
    Next m
    
    
    'Everything else is included
    IsTargetSheet = True

End Function

Function GetC2Value(ws As Worksheet) As String

    Dim result As String

    result = ws.Range("C2").value

    Select Case True

        'Template-Energy -> return blank
        Case InStr(1, ws.Name, "Template-Energy", vbTextCompare) > 0
            result = ""

        'QubeAU -> remove leading "C/O" then trim spaces
        Case InStr(1, ws.Name, "QubeAU", vbTextCompare) > 0

            If UCase(Left(Trim(result), 3)) = "C/O" Then
                result = Mid(Trim(result), 4)
            End If

            result = Trim(result)

        'TEPCO A -> append A
        Case InStr(1, ws.Name, "TEPCO A", vbTextCompare) > 0
            result = Trim(result) & " A"

        'TEPCO B -> append B
        Case InStr(1, ws.Name, "TEPCO B", vbTextCompare) > 0
            result = Trim(result) & " B"

    End Select

    GetC2Value = result

End Function


