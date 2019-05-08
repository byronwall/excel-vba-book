## SplitIntoRows.md

```vb
Public Sub SplitIntoRows()

    Dim outputRange As Range

    Dim inputRange As Range
    Set inputRange = Selection

    Set outputRange = GetInputOrSelection("Select the output corner")

    Dim targetPart As Variant
    Dim offsetCounter As Long
    offsetCounter = 0
    Dim targetCell As Range

    For Each targetCell In inputRange.SpecialCells(xlCellTypeVisible)
        Dim targetParts As Variant
        targetParts = Split(targetCell, vbLf)

        For Each targetPart In targetParts
            outputRange.Offset(offsetCounter) = targetPart

            offsetCounter = offsetCounter + 1
        Next targetPart
    Next targetCell
End Sub
```
