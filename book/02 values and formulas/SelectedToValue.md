## SelectedToValue.md

```vb
Public Sub SelectedToValue()

    Dim targetRange As Range
    On Error GoTo errHandler
    Set targetRange = GetInputOrSelection("Select the formulas you'd like to convert to static values")

    Dim targetCell As Range
    Dim targetCellValue As String
    For Each targetCell In targetRange
        targetCellValue = targetCell.Value
        targetCell.Clear
        targetCell = targetCellValue
    Next targetCell
    Exit Sub
errHandler:
    MsgBox "No selection made!"
End Sub
```