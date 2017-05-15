## RangeEnd_Boundary.md

```vb
Public Function RangeEnd_Boundary(ByVal rangeBegin As Range, ByVal firstDirection As XlDirection, Optional ByVal secondDirection As XlDirection = -1) As Range

    If secondDirection = -1 Then
        Set RangeEnd_Boundary = Intersect(Range(rangeBegin, rangeBegin.End(firstDirection)), rangeBegin.CurrentRegion)
    Else
        Set RangeEnd_Boundary = Intersect(Range(rangeBegin, rangeBegin.End(firstDirection).End(secondDirection)), rangeBegin.CurrentRegion)
    End If
End Function
```