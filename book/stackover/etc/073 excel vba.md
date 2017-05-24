# SO item 073
I seem to be getting an error on this, and I do not understand why, I know I can just use Range with letters but I want to learn how to do it in this format.

```
 ThisWorkbook.Sheets("t").Range(Cells(1, 1), Cells(2, 2)).Value = ThisWorkbook.Sheets("1").Range(Cells(1, 1), Cells(2, 2)).Value

```

----

The answer by @Sobigen gives a good way to qualify your references to avoid the error.

You can also avoid `Cells` altogether by using `Resize`.

```
Sub UseResizeInsteadOfCells()

    ThisWorkbook.Sheets("t").Range("A1").Resize(2, 2).Value = _
        ThisWorkbook.Sheets("1").Range("A1").Resize(2, 2).Value

End Sub

```

I used `A1` since you are doing `Cells(1,1)` on a `Worksheet` which is the same reference. You could also use `.Cells(1,1).Resize(2,2)` and get the same result without worrying about qualifying references inside a `Range` call.
