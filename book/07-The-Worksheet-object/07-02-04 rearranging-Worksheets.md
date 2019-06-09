### rearranging Worksheets

To rearrange the Worksheets, the command is simple: `Worksheet.Move(Before, After)`. The parametrs there will indicate the sheet to place it before or after. The real task here is determining which sheet to reference there, but finding that reference is the same task that is described up at the top of the section.

#### AscendSheets.md

TODO: move the AscendSheets code elsewhere or delete (not helpful here)

```vb
Public Sub AscendSheets()

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
                If UCase(.Sheets(i).name) > UCase(.Sheets(i + 1).name) Then .Sheets(i).Move after:=.Sheets(i + 1)
            Next i
        Next j
    End With

    Application.ScreenUpdating = True
End Sub
```
