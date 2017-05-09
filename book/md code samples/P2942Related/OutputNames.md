```vb
Public Sub OutputNames()

    Dim rngOut As Range
    Set rngOut = Range("B3")
    
    Dim namedRange As name
    For Each namedRange In ActiveWorkbook.Names
        
        Dim hasForm As Boolean
        hasForm = False
        
        On Error Resume Next
        hasForm = namedRange.RefersToRange.HasFormula
        On Error GoTo 0
        
        If namedRange.Visible And namedRange.name <> "SELF" Then
            
            rngOut = namedRange.name
            rngOut.Offset(, 1) = namedRange.Comment
            
            Set rngOut = rngOut.Offset(1)
            
        End If
    Next
End Sub
```