```vb
Sub Rand_ConvertToString()

    Dim cell As Range
    Dim sel As Range
    
    Set sel = Selection
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each cell In Intersect(sel, sel.Parent.UsedRange)
        If Not IsEmpty(cell.Value) And Not cell.HasFormula Then
            cell.Value = CStr(cell.Value)
        End If
    Next cell
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
```