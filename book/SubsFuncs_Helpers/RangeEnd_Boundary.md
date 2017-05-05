```vb
Public Function RangeEnd_Boundary(ByVal rangeBegin As Range, ByVal firstDirection As XlDirection, Optional ByVal secondDirection As XlDirection = -1) As Range
    '---------------------------------------------------------------------------------------
    ' Procedure : RangeEnd_Boundary
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Helper function to return a range limited by the starting cell's CurrentRegion
    '---------------------------------------------------------------------------------------
    '
    If secondDirection = -1 Then
        Set RangeEnd_Boundary = Intersect(Range(rangeBegin, rangeBegin.End(firstDirection)), rangeBegin.CurrentRegion)
    Else
        Set RangeEnd_Boundary = Intersect(Range(rangeBegin, rangeBegin.End(firstDirection).End(secondDirection)), rangeBegin.CurrentRegion)
    End If
End Function
```