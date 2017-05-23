# SO item 059
```
    For Each c In oSheet.Range("A1:A1000")
        If InStr(c.Value, "VALUE") Then
            c.EntireRow.Delete()
        End If
    Next

```

This will only delete a few of the rows within the specified range, what could the problem be?

----

Here are two common patterns for deleting entire rows based on a condition. The main idea is that you **cannot delete from a collection while you iterate it**. This means that `Delete` should not appear in a `For Each` This is fairly standard across most programming languages and some even throw an error to prevent it.

**Option 1**, use an integer to track the rows and have it work from the end to the beginning. You need to go backwards because it is the easy way to avoid skipping rows. It is possible to go forwards, you just need to account for not incrementing when you delete.

```
Sub DeleteRowsWithIntegerLoop()

    Dim rng_delete As Range
    Set rng_delete = Range("A1:A1000")

    Dim int_start As Integer
    int_start = rng_delete.Rows.Count

    Dim i As Integer
    For i = int_start To 1 Step -1
        If InStr(rng_delete.Cells(i), "VALUE") > 0 Then
            rng_delete.Cells(i).EntireRow.Delete
        End If
    Next i

End Sub

```

**Option 2**, use the `Union-Delete` pattern to build a range of cells and then delete them all in one step at the end.

```
Sub DeleteRowsWithUnionDelete()

    Dim rng_cell As Range
    Dim rng_delete As Range

    For Each rng_cell In Range("A1:A1000")
        If InStr(rng_cell, "VALUE") > 0 Then
            If rng_delete Is Nothing Then
                Set rng_delete = rng_cell
            Else
                Set rng_delete = Union(rng_delete, rng_cell)
            End If
        End If
    Next

    rng_delete.EntireRow.Delete

End Sub

```

**Notes on the code**

For Option 2, there is an extra conditional there to create `rng_delete` when it starts at the first item. `Union` does not work with a `Nothing` reference, so we first check that and if so, `Set` to the first item. All others come through and get `Set` by the `Union` line.

**Preference**

When choosing between the two, I always prefer Option 2 because I much prefer to work with `Ranges` in Excel instead of iterating through `Cells` with a counter. There are limitations to this. The second option also works for discontinuous `Ranges` and all other variety of weird `Ranges` (e.g. after a call to `SpecialCells`) which can make it valuable when you are not sure what data you will be dealing with.

**Speed**

I am not sure about a speed comparison. Both can be slow if `ScreenUpdating` and calculations are enabled. The first option makes `N-1` calls to `Delete` whereas the second option does a single one. `Delete` is an expensive operation. Option 2 does however make `N-1` calls to `Union` and `Set`. I assume it's faster than the first one based on that (and it seems to be here), but I did not profile it.

Final note: `InStr` returns an integer indicating where the value was found. I always like to make the boolean covnersion explicit here and compare to `>0`.
