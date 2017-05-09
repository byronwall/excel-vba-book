```vb
Public Sub SelectedToValue()
    '---------------------------------------------------------------------------------------
    ' Procedure : SelectedToValue
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Forces a targetCell to take on its value.  Removes formulas.
    '---------------------------------------------------------------------------------------
    '
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