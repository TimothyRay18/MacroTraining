Attribute VB_Name = "OHStock"
Function getMaxRow(col As Integer) As Double
    getMaxRow = ActiveSheet.Cells(Rows.Count, col).End(xlUp).row
End Function

Function getMaxCol(row As Integer) As Double
    getMaxCol = ActiveSheet.Cells(row, Columns.Count).End(xlToLeft).Column
End Function

Function findCellInColumn(row As Integer, str As String) As Double
    Dim i As Double
    i = 1
    Dim m As Double
    m = getMaxCol(row)
    While LCase(ActiveSheet.Cells(row, i).Value) <> LCase(str) And i <= m
        i = i + 1
    Wend
    findCellInColumn = i
End Function

Sub OH()
    Dim oh_file As String
    oh_file = Range("B3").Value
    Workbooks.Open Filename:=oh_file
    
    Dim max_col As Double
    max_col = getMaxCol(1)
    Dim max_row As Double
    max_row = getMaxRow(1)
    
    'Range("L1").Select
    Cells(1, max_col + 1).Select
    ActiveCell.FormulaR1C1 = "Ok"
    'Range("M1").Select
    Cells(1, max_col + 2).Select
    ActiveCell.FormulaR1C1 = "Sloc"
    
    Cells(2, max_col + 1).Select
    ActiveCell.FormulaR1C1 = "=RC[-8]+RC[-7]+RC[-6]"
    Selection.AutoFill Destination:=Range(Cells(2, max_col + 1), Cells(max_row, max_col + 1))
    
    Cells(2, max_col + 2).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[-11],1)=""L"",""Prod"",IF(RC[-11]=""0012"",""LTB"",IF(LEFT(RC[-11],1)=""9"",""Quarantine"",IF(RC[-11]="""",""SubCon"",""WH""))))"
    Selection.AutoFill Destination:=Range(Cells(2, max_col + 2), Cells(max_row, max_col + 2))
    
    Dim source As String
    source = "Sheet1!R1C1:R" + CStr(max_row) + "C" + CStr(max_col + 2)
    Sheets.Add.Name = "PivotTable"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        source, Version:=6).CreatePivotTable TableDestination:= _
        "PivotTable!R3C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("PivotTable").Select
    
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Material")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Sloc")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Ok"), "Sum of Ok", xlSum
End Sub
