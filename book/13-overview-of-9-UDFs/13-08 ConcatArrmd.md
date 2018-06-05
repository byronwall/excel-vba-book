## ConcatArr.md

```vb
Public Function ConcatArr(rngCells As Variant, strDelim As String) As String
    Dim cellCount As Long
    
    cellCount = UBound(rngCells, 1)
    
    Dim arrValues As Variant
    ReDim arrValues(1 To cellCount)
    
    Dim index As Long
    index = 1
    
    Dim rngCell As Variant
    For Each rngCell In rngCells
        arrValues(index) = rngCell
        
        index = index + 1
    Next
    
    ConcatArr = Join(arrValues, strDelim)
End Function
```