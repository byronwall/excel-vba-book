```vb
Public Sub ChangeNameIfInRange()

    Dim rngOut As Range
    Set rngOut = Selection
    
    Dim namedRange As name
    For Each namedRange In ActiveWorkbook.Names
        
        Dim hasForm As Boolean
        hasForm = False
        
        On Error Resume Next
        Dim rngName As Range
        Set rngName = namedRange.RefersToRange
        On Error GoTo 0
        
        If Not rngName Is Nothing Then
            If Not Intersect(rngName, rngOut) Is Nothing Then
            
                Debug.Print namedRange.name
                namedRange.name = Replace(namedRange.name, "mole", "sl")
                Debug.Print namedRange.name
                Debug.Print ""
            End If
        
        End If
        
        
    Next
    

End Sub
```