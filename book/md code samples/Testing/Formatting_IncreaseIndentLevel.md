```vb
Sub Formatting_IncreaseIndentLevel()

    Dim rngCell As Range
    
    For Each rngCell In Selection
        rngCell.IndentLevel = rngCell.IndentLevel + 2
    Next

End Sub
```