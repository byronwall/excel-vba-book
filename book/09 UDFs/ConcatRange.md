## ConcatRange.md

```vb
Public Function ConcatRange(rngCells As Range, strDelim As String) As String
    Dim cellCount As Long
    
    cellCount = rngCells.CountLarge
    
    Dim arrValues As Variant
    ReDim arrValues(1 To cellCount)
    
    Dim index As Long
    index = 1
    
    Dim rngCell As Range
    For Each rngCell In rngCells
        arrValues(index) = rngCell
        
        index = index + 1
    Next
    
    ConcatRange = Join(arrValues, strDelim)
End Function
```