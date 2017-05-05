```vb
Sub Rand_DumpTextFromAllSheets()

    Dim c As Range
    Dim s As Worksheet
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim main As Workbook
    Set main = ActiveWorkbook
    
    Dim w As Workbook
    Dim sw As Worksheet
    
    Set w = Application.Workbooks.Add
    Set sw = w.Sheets.Add
    
    Dim Row As Long
    Row = 0
    For Each s In main.Sheets
        For Each c In s.UsedRange.SpecialCells(xlCellTypeConstants)
            sw.Range("A1").Offset(Row) = c
            Row = Row + 1
        Next c
    Next s

End Sub
```