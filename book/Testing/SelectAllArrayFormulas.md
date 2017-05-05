```vb
Public Sub SelectAllArrayFormulas()
    '---------------------------------------------------------------------------------------
    ' Procedure : SelectAllArrayFormulas
    ' Author    : @byronwall
    ' Date      : 2016 01 27
    ' Purpose   : selects all cells on current sheet that have an array formula
    '---------------------------------------------------------------------------------------
    '

    Dim rng_forms As Range

    Set rng_forms = ActiveSheet.UsedRange

    Dim rng_select As Range

    Dim rng_form As Range
    For Each rng_form In rng_forms
        If rng_form.HasArray Then
            If rng_select Is Nothing Then
                Set rng_select = rng_form
            Else
                Set rng_select = Union(rng_select, rng_form)
            End If
        End If
    Next

    rng_select.Select

End Sub
```