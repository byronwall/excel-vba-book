```vb
Sub TraceDependentsForAll()
    '---------------------------------------------------------------------------------------
    ' Procedure : TraceDependentsForAll
    ' Author    : @byronwall
    ' Date      : 2015 11 09
    ' Purpose   : Quick Sub to iterate through Selection and Trace Dependents for all
    '---------------------------------------------------------------------------------------
    '
    Dim rng As Range
    
    For Each rng In Intersect(Selection, Selection.Parent.UsedRange)
        rng.ShowDependents
    Next rng

End Sub
```