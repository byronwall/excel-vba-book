```vb
Sub ApplyFormattingToEachColumn()
    Dim rng As Range
    For Each rng In Selection.Columns

        Formatting_AddCondFormat rng
    Next
End Sub
```