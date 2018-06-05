```vb
Public Sub OutputFontSizeNextToCell()

    'this is a quick routine to output the font size next to a cell
    Dim rngAllCells As Range
    Set rngAllCells = GetInputOrSelection("Select range")
    
    Dim rngCell As Range
    For Each rngCell In rngAllCells
        If rngCell <> "" Then
            rngCell.Offset(, 1) = rngCell.Font.Size
        End If
    Next
    
    
End Sub
```