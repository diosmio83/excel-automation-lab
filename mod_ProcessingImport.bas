Attribute VB_Name = "Modul1"
Sub RunStepsThenImport()
    Call SortHeadersWithDataRealignment
    Call ApplyFormatting
    Call CopyColumnLToMAndN
    Call SortByColumnH
    Call FillFixedValues
    Call FormatColumnLAsText
    Call ImportProcessingData_FixedColumns_LMN_ABFO
    Call FormatColumnLAsNumber
End Sub

' Schritt 1: Header alphabetisch sortieren und Daten korrekt zuordnen
Sub SortHeadersWithDataRealignment()
    Dim ws As Worksheet, tempWs As Worksheet
    Dim headerRange As Range, dataRange As Range
    Dim headers As Variant, data As Variant
    Dim colCount As Long, rowCount As Long

    Set ws = ActiveSheet
    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    rowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(rowCount, colCount))

    headers = headerRange.Value
    data = dataRange.Value

    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("__TempSort").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set tempWs = Worksheets.Add
    tempWs.Name = "__TempSort"

    tempWs.Range(tempWs.Cells(1, 1), tempWs.Cells(1, colCount)).Value = headers
    tempWs.Range(tempWs.Cells(2, 1), tempWs.Cells(1 + UBound(data), colCount)).Value = data

    With tempWs.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=tempWs.Range(tempWs.Cells(1, 1), tempWs.Cells(1, colCount)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange tempWs.Range(tempWs.Cells(1, 1), tempWs.Cells(1 + UBound(data), colCount))
        .Header = xlYes
        .Orientation = xlLeftToRight
        .Apply
    End With

    headerRange.Value = tempWs.Range(tempWs.Cells(1, 1), tempWs.Cells(1, colCount)).Value
    dataRange.Value = tempWs.Range(tempWs.Cells(2, 1), tempWs.Cells(1 + UBound(data), colCount)).Value

    Application.DisplayAlerts = False
    tempWs.Delete
    Application.DisplayAlerts = True
End Sub

Sub ApplyFormatting()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    On Error Resume Next
    ws.Range("D2:D1048576").NumberFormat = "dd-mmm-yyyy"
    ws.Range("E2:E1048576").NumberFormat = "hh:mm;@"
    ws.Range("L2:L1048576").NumberFormat = "0"
    ws.Range("M2:M1048576").NumberFormat = "0"
    ws.Range("N2:N1048576").NumberFormat = "0"
    ws.Range("V2:V1048576").NumberFormat = "0"
    On Error GoTo 0
End Sub

Sub CopyColumnLToMAndN()
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row

    ws.Range("M2:M" & lastRow).NumberFormat = "0"
    ws.Range("N2:N" & lastRow).NumberFormat = "0"

    ws.Range("M2:M" & lastRow).Value = ws.Range("L2:L" & lastRow).Value
    ws.Range("N2:N" & lastRow).Value = ws.Range("L2:L" & lastRow).Value
End Sub

Sub SortByColumnH()
    Dim ws As Worksheet
    Dim lastRow As Long, colCount As Long
    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("H2:H" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    With ws.Sort
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, colCount))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub FillFixedValues()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim colQ As Long, colG As Long, colI As Long

    Set ws = ActiveSheet

    On Error Resume Next
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0

    colQ = GetColumnIndexByHeader(ws, "Parent Volume Unit")
    colG = GetColumnIndexByHeader(ws, "Aliquot Volume Unit")
    colI = GetColumnIndexByHeader(ws, "HIV Status")

    If colQ > 0 Then ws.Range(ws.Cells(2, colQ), ws.Cells(lastRow, colQ)).Value = "mL"
    If colG > 0 Then ws.Range(ws.Cells(2, colG), ws.Cells(lastRow, colG)).Value = "uL"
    If colI > 0 Then ws.Range(ws.Cells(2, colI), ws.Cells(lastRow, colI)).Value = "HIV inactivated"
End Sub

Function GetColumnIndexByHeader(ws As Worksheet, headerName As String) As Long
    Dim cell As Range
    Dim lastCol As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For Each cell In ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        If Trim(UCase(cell.Value)) = UCase(headerName) Then
            GetColumnIndexByHeader = cell.Column
            Exit Function
        End If
    Next cell

    GetColumnIndexByHeader = 0
End Function

Sub FormatColumnLAsText()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Columns(12).NumberFormat = "@"
End Sub

Sub FormatColumnLAsNumber()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Columns(12).NumberFormat = "0"
End Sub

Function FormatColumnBValue(ByVal inputValue As Variant) As String
    Dim sValue As String
    sValue = CStr(inputValue)
    If IsNumeric(sValue) And Len(sValue) < 4 Then
        FormatColumnBValue = Right$("0000" & sValue, 4)
    Else
        FormatColumnBValue = sValue
    End If
End Function

Function SelectExcelFile() As String
    Dim filePath As Variant
    filePath = Application.GetOpenFilename("Excel-Dateien (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", , "Please select processing file")
    If filePath <> False Then
        SelectExcelFile = filePath
    Else
        SelectExcelFile = ""
    End If
End Function

Sub ImportProcessingData_FixedColumns_LMN_ABFO()
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim filePath As String
    Dim lastRowSource As Long, lastRowTarget As Long
    Dim i As Long, j As Long
    Dim startTime As Date
    Dim matchCount As Integer

    Set targetWs = ActiveSheet
    filePath = SelectExcelFile()
    If filePath = "" Then Exit Sub

    startTime = Now
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceWs = sourceWb.Sheets(1)

    lastRowSource = sourceWs.Cells(sourceWs.Rows.Count, "A").End(xlUp).Row
    lastRowTarget = targetWs.Cells(targetWs.Rows.Count, "L").End(xlUp).Row

    targetWs.Columns("B").NumberFormat = "@"

    Dim logWs As Worksheet
    On Error Resume Next
    Set logWs = ThisWorkbook.Worksheets("ImportLog")
    If logWs Is Nothing Then
        Set logWs = ThisWorkbook.Worksheets.Add
        logWs.Name = "ImportLog"
    End If
    On Error GoTo 0
    logWs.Cells.ClearContents
    logWs.Cells(1, 1).Value = "Not found IDs"
    logWs.Cells(1, 2).Value = "Target Row"
    Dim logRow As Long: logRow = 2

    Dim notFoundCount As Long: notFoundCount = 0
    Dim successCount As Long: successCount = 0
    Dim totalRows As Long: totalRows = lastRowTarget - 1

    For i = 2 To lastRowTarget
        Dim idTarget As String
        idTarget = Trim(CStr(targetWs.Cells(i, 12).Value))

        If idTarget <> "" Then
            Dim found As Boolean: found = False
            matchCount = 0
            For j = 2 To lastRowSource
                If Trim(CStr(sourceWs.Cells(j, 9).Value)) = idTarget Then
                    matchCount = matchCount + 1
                End If
            Next j
            If matchCount > 1 Then
                targetWs.Rows(i).Interior.Color = RGB(255, 255, 150)
            End If
            For j = 2 To lastRowSource
                If Trim(CStr(sourceWs.Cells(j, 9).Value)) = idTarget Then
                    targetWs.Cells(i, 1).Value = sourceWs.Cells(j, 17).Value
                    targetWs.Cells(i, 2).Value = FormatColumnBValue(sourceWs.Cells(j, 4).Value)
                    targetWs.Cells(i, 6).Value = sourceWs.Cells(j, 5).Value
                    targetWs.Cells(i, 15).Value = sourceWs.Cells(j, 10).Value
                    targetWs.Rows(i).Interior.Color = RGB(230, 255, 230)
                    found = True
                    successCount = successCount + 1
                    Exit For
                End If
            Next j
            If Not found Then
                notFoundCount = notFoundCount + 1
                targetWs.Rows(i).Interior.Color = RGB(255, 230, 230)
                logWs.Cells(logRow, 1).Value = idTarget
                logWs.Cells(logRow, 2).Value = i
                logRow = logRow + 1
            End If
        End If
    Next i

    sourceWb.Close False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Dim endTime As Date: endTime = Now
    Dim duration As Double: duration = (endTime - startTime) * 86400
    Dim successRate As Double: successRate = (successCount / totalRows) * 100

    MsgBox "Import Results:" & vbCrLf & _
           "? Successfully imported: " & successCount & " of " & totalRows & " rows (" & Format(successRate, "0.0") & "%)" & vbCrLf & _
           "? Not found: " & notFoundCount & " identifiers" & vbCrLf & _
           "?? Details in sheet 'ImportLog'" & vbCrLf & _
           "? Duration: " & Format(duration, "0.0") & " seconds", vbInformation
End Sub
    ChDir "C:\Users\DiosMio\Downloads"
    ActiveWorkbook.SaveAs Filename:="C:\Users\DiosMio\Downloads\Mappe1makro.xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    ActiveWorkbook.Save

