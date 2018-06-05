## FillValueDown.md

```vb
Public Sub FillValueDown()

    Dim inputRange As Range
    Set inputRange = GetInputOrSelection("Select range for waterfall")

    If inputRange Is Nothing Then Exit Sub

    Dim targetCell As Range
    For Each targetCell In Intersect(inputRange.SpecialCells(xlCellTypeBlanks), inputRange.Parent.UsedRange)
        targetCell = targetCell.End(xlUp)
    Next targetCell

End Sub
```