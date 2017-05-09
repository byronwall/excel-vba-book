```vb
Public Sub SetFormulaToNeighborCell()

    Dim rngCell As Range
    
    For Each rngCell In Selection
        Dim strForm As String
        strForm = rngCell.Offset(, -1).Value
        
        If strForm <> "" Then
            rngCell.Formula = "=" & strForm
        End If
    Next

End Sub
```