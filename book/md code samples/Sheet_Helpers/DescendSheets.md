```vb
Public Sub DescendSheets()
    '---------------------------------------------------------------------------------------
    ' Procedure : DescendSheets
    ' Author    : @raymondwise
    ' Date      : 2015 08 07
    ' Purpose   : Places worksheets in descending alphabetical order.
    '---------------------------------------------------------------------------------------
    Application.ScreenUpdating = False
    Dim targetWorkbook As Workbook
    Set targetWorkbook = ActiveWorkbook

    Dim countOfSheets As Long
    countOfSheets = targetWorkbook.Sheets.Count

    Dim i As Long
    Dim j As Long

    With targetWorkbook
        For j = 1 To countOfSheets
            For i = 1 To countOfSheets - 1
                If UCase(.Sheets(i).name) < UCase(.Sheets(i + 1).name) Then .Sheets(i).Move after:=.Sheets(i + 1)
            Next i
        Next j
    End With

    Application.ScreenUpdating = True
End Sub
```