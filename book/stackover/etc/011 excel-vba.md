# SO item 011
I want to get a new range by doing the subtract of two other existing ranges. Let's say I have : `Set rng1=Range("A1")` and `Set rng2=Range("A1:A5")` I want to calculate a new range : rng3 = rng2 - rng1. I've tried : `Set rng3 = minus(rng2, rng1)` but it seems it's not possible.

----

Here is one approach with a UDF. It may not be particularly fast on large ranges since it iterates cell by cell. I suspect it will handle most cases though.

```
Public Function DisUnion(keep As Range, remove As Range) As Range

    Dim rng_output As Range

    Dim cell As Range
    For Each cell In keep

        'check if given cell is in range to remove
        If Intersect(cell, remove) Is Nothing Then

            'this builds the output and handles first case
            If rng_output Is Nothing Then
                Set rng_output = cell
            Else
                Set rng_output = Union(rng_output, cell)
            End If
        End If
    Next cell

    Set DisUnion = rng_output

End Function

```

# Usage

The result of the cell below is 33 which is correct. It updates to changes to cells as expected.

![enter image description here](https://i.stack.imgur.com/93KxA.png)
