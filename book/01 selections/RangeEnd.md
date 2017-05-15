## RangeEnd.md

```vb
Public Function RangeEnd(ByVal rangeBegin As Range, ByVal firstDirection As XlDirection, Optional ByVal secondDirection As XlDirection = -1) As Range

    If secondDirection = -1 Then
        Set RangeEnd = Range(rangeBegin, rangeBegin.End(firstDirection))
    Else
        Set RangeEnd = Range(rangeBegin, rangeBegin.End(firstDirection).End(secondDirection))
    End If
End Function
```