```vb
Public Function RangeEnd(ByVal rangeBegin As Range, ByVal firstDirection As XlDirection, Optional ByVal secondDirection As XlDirection = -1) As Range
    '---------------------------------------------------------------------------------------
    ' Procedure : RangeEnd
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Helper function to return a block of cells using a starting Range and an End firstDirection
    '---------------------------------------------------------------------------------------
    '
    If secondDirection = -1 Then
        Set RangeEnd = Range(rangeBegin, rangeBegin.End(firstDirection))
    Else
        Set RangeEnd = Range(rangeBegin, rangeBegin.End(firstDirection).End(secondDirection))
    End If
End Function
```