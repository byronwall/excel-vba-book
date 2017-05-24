# SO item 009
Is there a more efficient way to do the below operation of adding three cell values in a row to the corresponding cells above and then deleting the old row? Half the time the macro freezes; I am running it on about 12,000 lines (there are no dynamic formulas on the sheet).

```
Application.ScreenUpdating = False

For a = Lcell To 2 Step -1
   If Cells(a, 23).Value = Cells(a - 1, 23).Value Then
       Cells(a, 16).Value = Cells(a, 16).Value + Cells(a - 1, 16).Value
       Cells(a, 17).Value = Cells(a, 17).Value + Cells(a - 1, 17).Value
       Cells(a, 18).Value = Cells(a, 18).Value + Cells(a - 1, 18).Value
       Cells(a - 1, 1).EntireRow.Delete
   End If
Next a

Application.ScreenUpdating = True

```

----

One option when deleting cells is to use the UNION-DELETE pattern. This saves the deletion step until after the logic is determined and does it all at once. This technique allows deleting a Range that is being iterated through. It also reduces operations which should improve increase speed. I have not tested it for speed though.

**Edited code to delete last row based on comments**

```
Dim rng_delete As Range

For A = Lcell To 2 Step -1
   If Cells(A, 23).Value = Cells(A - 1, 23).Value Then
       Cells(A - 1, 16).Value = Cells(A, 16).Value + Cells(A - 1, 16).Value
       Cells(A - 1, 17).Value = Cells(A, 17).Value + Cells(A - 1, 17).Value
       Cells(A - 1, 18).Value = Cells(A, 18).Value + Cells(A - 1, 18).Value

        'rng_delete starts empty which errors Union on first add
        If rng_delete Is Nothing Then
            Set rng_delete = Cells(A, 1).EntireRow
        Else
            Set rng_delete = Union(rng_delete, Cells(A, 1).EntireRow)
        End If
   End If
Next A

rng_delete.Delete

```
