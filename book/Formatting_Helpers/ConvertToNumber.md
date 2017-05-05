```vb
Public Sub ConvertToNumber()
    '---------------------------------------------------------------------------------------
    ' Procedure : ConvertToNumber
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Forces all numbers stored as text to be converted to actual numbers
    '---------------------------------------------------------------------------------------
    '
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