## creating and managing Worksheets

This section will focus on how to create a Worksheet and get a reference to new Worksheets.  In addition to that, it will discuss managing Worksheets, including rearranging and deleting them.

### creating a Worksheet

TODO: add this content, focus on how to get a reference to the newly created sheet, also focus on how to control where the new sheet is place

TODO: also include a section about how to Copy an existing Worksheet

### removing a Worksheet

TODO: add this content about removing a Worksheet, focus on disabling the alerts that might show up

### rearranging Worksheets

TODO: add content about how to reorder the Worksheets

#### AscendSheets.md

TODO: clean up this code

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