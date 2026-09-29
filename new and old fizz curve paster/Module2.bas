Attribute VB_Name = "Module2"
Option Explicit

'============================================================
' Macro: Import New Japan Fizz Curve (Base / Peak Combined)
'============================================================
Public Sub New_Japan_Fizz_Curve()

    Dim wbOrigin As Workbook
    Dim wbDest As Workbook
    Dim wsCurve As Worksheet
    Dim wsMarks As Worksheet

    Dim originPattern As String
    Dim destPattern As String
    Dim todayYYMMDD As String
    Dim todayDDMMYY As Date

    Dim wb As Workbook
    Dim f As Range
    Dim x As Range
    
    Dim headerRow As Long

    '--------------------------------------------------------
    ' Date/User input
    '--------------------------------------------------------
    todayYYMMDD = Format(Sheet1.Range("A3").value, "yy.mm.dd")
    todayDDMMYY = Sheet1.Range("A3").value

    'Getting header row from user input in A5
    headerRow = CLng(ActiveSheet.Range("A5").value)
    
    '--------------------------------------------------------
    ' Workbook patterns
    '--------------------------------------------------------
    originPattern = "*FIZZ CURVE SHEET - MASTER v1*"
    destPattern = "*Vanir Japan Power Curve_PHYSICAL_" & todayYYMMDD & " NEW FORMAT*"

    '--------------------------------------------------------
    ' Find origin workbook
    '--------------------------------------------------------
    For Each wb In Workbooks
        If wb.Name Like originPattern Then
            Set wbOrigin = wb
            Exit For
        End If
    Next wb

    If wbOrigin Is Nothing Then
        MsgBox "Origin workbook not open", vbCritical
        Exit Sub
    End If

    '--------------------------------------------------------
    ' Find destination workbook
    '--------------------------------------------------------
    For Each wb In Workbooks
        If wb.Name Like destPattern Then
            Set wbDest = wb
            Exit For
        End If
    Next wb

    If wbDest Is Nothing Then
        MsgBox "Destination workbook not open", vbCritical
        Exit Sub
    End If

    '--------------------------------------------------------
    ' Resolve sheets
    '--------------------------------------------------------
    Set wsCurve = GetSheetByNameInsensitive(wbOrigin, "Base_Peak_Combined")
    Set wsMarks = GetSheetByNameInsensitive(wbDest, "MARKS")

    If wsCurve Is Nothing Or wsMarks Is Nothing Then
        MsgBox "Required sheet missing", vbCritical
        Exit Sub
    End If

    '--------------------------------------------------------
    ' Regions (order matters)
    '--------------------------------------------------------
    Dim regions As Variant
    regions = Array("Tokyo", "Chubu", "Kansai", "Hokkaido", "Tohoku", _
                    "Hokuriku", "Chugoku", "Shikoku", "Kyushu")

    '--------------------------------------------------------
    ' Paste sequence
    '--------------------------------------------------------
    Dim pasteOrder As Variant
    pasteOrder = Array( _
        "Tokyo|Base", "Tokyo|Peak", _
        "Chubu|Base", "Chubu|Peak", _
        "Kansai|Base", "Kansai|Peak", _
        "Hokkaido|Base", "Hokkaido|Peak", _
        "Tohoku|Base", "Tohoku|Peak", _
        "Hokuriku|Base", "Hokuriku|Peak", _
        "Chugoku|Base", "Chugoku|Peak", _
        "Shikoku|Base", "Shikoku|Peak", _
        "Kyushu|Base", "Kyushu|Peak" _
    )
    
    

    If headerRow <= 0 Then
        MsgBox "Invalid header row in A5", vbCritical
        Exit Sub
    End If


    '--------------------------------------------------------
    ' Locate region headers
    '--------------------------------------------------------
    Dim regionCol As Object
    Set regionCol = CreateObject("Scripting.Dictionary")

    Dim r As Variant
    For Each r In regions
        Set f = wsCurve.Rows(headerRow).Find( _
            What:=r, _
            LookAt:=xlWhole, _
            MatchCase:=False _
        )

        If f Is Nothing Then
            MsgBox "Region header not found: " & r, vbCritical
            Exit Sub
        End If

        regionCol(r) = f.Column
    Next r

    '--------------------------------------------------------
    ' Determine BASE range from Tokyo Base
    '--------------------------------------------------------
    Dim baseRowTokyo As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim colContract As Long

    Set f = wsCurve.Columns(regionCol("Tokyo")).Find( _
        What:="Base", _
        LookAt:=xlWhole, _
        MatchCase:=False _
    )

    If f Is Nothing Then
        MsgBox "Tokyo Base not found", vbCritical
        Exit Sub
    End If

    baseRowTokyo = f.Row
    firstDataRow = baseRowTokyo + 2
    colContract = regionCol("Tokyo") - 1

    lastDataRow = wsCurve.Cells(wsCurve.Rows.Count, colContract).End(xlUp).Row

    If lastDataRow < firstDataRow Then
        MsgBox "No data found under Tokyo Base", vbCritical
        Exit Sub
    End If

    '--------------------------------------------------------
    ' Clear destination (keep headers)
    '--------------------------------------------------------
    Dim destRow As Long
    Dim lastRow As Long

    lastRow = wsMarks.Cells(wsMarks.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then wsMarks.Rows("2:" & lastRow).ClearContents

    destRow = 2
    
    ' Set formats at the top, before the loop
    With wsMarks
        .Columns(1).NumberFormat = "dd mmm yyyy"  ' TODAY column as Date
        .Columns(2).NumberFormat = "General"
        .Columns(3).NumberFormat = "[$-en-US]mmm-yy"      ' CONTRACT column as Date with custom display
        .Columns(4).NumberFormat = "0.00"        ' Marks
    End With

    'wsMarks.Columns(3).NumberFormat = "@"      ' Contract (text)
    'wsMarks.Columns(4).NumberFormat = "@"      ' Marks
    'wsMarks.Columns(5).NumberFormat = "@"      ' Change

    '--------------------------------------------------------
    ' Main loop
    '--------------------------------------------------------
    Dim item As Variant
    Dim parts() As String
    Dim regionName As String
    Dim productType As String
    Dim dataCol As Long
    Dim markVal As Variant
    Dim contractVal As Variant
    Dim rowPtr As Long
    Dim firstDataRowLocal As Long ' <-- local firstDataRow per region/product
    
    For Each item In pasteOrder
    
        parts = Split(item, "|")
        regionName = Trim$(parts(0))
        productType = Trim$(parts(1))
    
        ' Find Base / Peak row for region
       Set x = wsCurve.Rows(headerRow + 2).Find( _
            What:=productType, _
            LookAt:=xlWhole, _
            MatchCase:=False _
        )
    
        If x Is Nothing Then
            MsgBox productType & " not found for " & regionName, vbCritical
            Exit Sub
        End If
    
        dataCol = regionCol(regionName)
        
        firstDataRowLocal = x.Row + 1 ' Data starts after Base/Peak row
        
        
        ' If it's Peak, move one column to the right
        If productType = "Peak" Then dataCol = dataCol + 1
        
        firstDataRowLocal = baseRowTokyo + 2  ' or whatever row your data starts

    
        ' Copy data from firstDataRowLocal to lastDataRow
        ' Copy data from firstDataRowLocal to lastDataRow
        For rowPtr = firstDataRowLocal To lastDataRow
    
            contractVal = wsCurve.Cells(rowPtr, colContract).value
            markVal = wsCurve.Cells(rowPtr, dataCol).value
    
            wsMarks.Cells(destRow, 1).value = todayDDMMYY
            wsMarks.Cells(destRow, 2).value = regionName & " Area " & productType & "load"
            ' If it is a real date, force to 1st of month
            If IsDate(contractVal) Then
                wsMarks.Cells(destRow, 3).value = _
                DateSerial(Year(contractVal), Month(contractVal), 1)
            Else
                wsMarks.Cells(destRow, 3).value = contractVal
            End If


            wsMarks.Cells(destRow, 4).value = markVal
    
            destRow = destRow + 1
        Next rowPtr
    
    Next item


    '-------------------------------
    ' Format Header Row (A1:D1)
    '-------------------------------
    With wsMarks.Range("A1:D1")
        .Font.Bold = True                  ' Bold
        .Borders.LineStyle = xlContinuous  ' All borders
    End With
    
    '-------------------------------
    ' Enable AutoFilter
    '-------------------------------
    If wsMarks.AutoFilterMode Then wsMarks.AutoFilterMode = False
    wsMarks.Range("A1:E1").AutoFilter
    
    '-------------------------------
    ' Clear Any Active Filters
    '-------------------------------
    If wsMarks.FilterMode Then wsMarks.ShowAllData

    wbDest.Save
    MsgBox "Done", vbInformation

End Sub

'============================================================
' Helper: Get Sheet by Name (Case-Insensitive)
'============================================================
Public Function GetSheetByNameInsensitive(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(Trim(ws.Name), Trim(sheetName), vbTextCompare) = 0 Then
            Set GetSheetByNameInsensitive = ws
            Exit Function
        End If
    Next ws
End Function


