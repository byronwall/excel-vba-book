# SO item 041
I have target.address (or just target) from a worksheet_change sub. I'd like to use the row from target.address and a range of columns (H:W), and get a sum of that range. So, for instance, if I have $100 in H10 and I add $50 in J10, I'd like to get the sum of $150 since my target.address row is 10 and I'm within my desired column range H:W.

----

You can use `Intersect` and `EntireRow` to figure out which cells to sum. I would give them to `Application.Sum` to do the math. Another call to `Intersect` will let you know if the changed cell is in the "boxed" area.

```
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim sum As Double
    Dim rng_match As Range

    Set rng_match = Range("H:W")

    If Target.Cells.Count = 1 Then
        If Not Intersect(Target, rng_match) Is Nothing Then
            sum = Application.sum(Intersect(Target.EntireRow, rng_match))
        End If
    End If

End Sub

```

**Edit:** added a check on `Target.Cells.Count` per @Tim Williams to ensure that the sum is not affected by a multi-cell edit. The `Else` would need to address what to do next if this is an issue.
