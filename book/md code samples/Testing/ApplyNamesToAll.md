```vb
Public Sub ApplyNamesToAll()

    
        
    Dim nameCount As Long
    nameCount = 0
    
    Dim rngName As name
    For Each rngName In ActiveWorkbook.Names
        If rngName.Visible Then
            nameCount = nameCount + 1
        End If
    Next
    
    Dim arrNames() As String
    ReDim arrNames(1 To nameCount)
    
    Dim counter As Long
    counter = 1
    
    For Each rngName In ActiveWorkbook.Names
        If rngName.Visible Then
            
            arrNames(counter) = rngName.name
            counter = counter + 1
        End If
        
    Next
    
    Dim rng As Range
    Set rng = Selection
    
    rng.ApplyNames arrNames, True, False
    

End Sub
```