```vb
Public Sub FormulaModifier()

    'this works for the single case where formula is =A+B
    'it will substitute the constituent formulas for A & B
    'this does not work in the general case at all

    'get the current formula
    Dim rngCell As Range
    For Each rngCell In Selection
    
        'remove the first =
        Dim strForm As String
        strForm = rngCell.Formula
    
        strForm = Right(strForm, Len(strForm) - 1)
    
        'split based on + sign
        Dim parts As Variant
        parts = Split(strForm, "+")
    
        Dim newParts() As String
        ReDim newParts(UBound(parts))
    
        Dim index As Long
        For index = LBound(parts) To UBound(parts)
            Dim strPartForm As String
            strPartForm = Range(parts(index)).Formula
            newParts(index) = Right(strPartForm, Len(strPartForm) - 1)
        Next
    
        Dim strNewForm As String
        strNewForm = "=" & Join(newParts, "+")
    
        'get the cells and parse their formulas
    
        rngCell.Formula = strNewForm
    Next
    
    'stick those formulas into the current one

End Sub
```