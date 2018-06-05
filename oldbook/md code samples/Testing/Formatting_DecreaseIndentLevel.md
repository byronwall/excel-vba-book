```vb
Sub Formatting_DecreaseIndentLevel()

    Dim rngCell As Range
    
    For Each rngCell In Selection
        rngCell.IndentLevel = WorksheetFunction.Max(rngCell.IndentLevel - 2, 0)
    Next

End Sub
```