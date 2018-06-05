## ConvertToNumber.md

```vb
Public Sub ConvertToNumber()

    Dim targetCell As Range
    Dim targetSelection As Range

    Set targetSelection = Selection

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each targetCell In Intersect(targetSelection, ActiveSheet.UsedRange)
        If Not IsEmpty(targetCell.Value) And IsNumeric(targetCell.Value) Then targetCell.Value = CDbl(targetCell.Value)
    Next targetCell

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
```