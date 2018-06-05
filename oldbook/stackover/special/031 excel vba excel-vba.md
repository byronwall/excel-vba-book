# SO item 031
**I have the following:**

![enter image description here](https://i.stack.imgur.com/yTqt9.png)

**I expect the following:**

![enter image description here](https://i.stack.imgur.com/neZ3O.png)

**I am using this code:**

```
Sub merge_cells()

Application.DisplayAlerts = False

Dim r As Integer
Dim mRng As Range
Dim rngArray(1 To 4) As Range
r = Range("A65536").End(xlUp).Row

For myRow = r To 2 Step -1

    If Range("A" & myRow).Value = Range("A" & (myRow - 1)).Value Then

        For cRow = (myRow - 1) To 1 Step -1

            If Range("A" & myRow).Value <> Range("A" & cRow).Value Then

                Set rngArray(1) = Range("A" & myRow & ":A" & (cRow + 0))
                Set rngArray(2) = Range("B" & myRow & ":B" & (cRow + 0))
                Set rngArray(3) = Range("C" & myRow & ":C" & (cRow + 0))
                Set rngArray(4) = Range("D" & myRow & ":D" & (cRow + 0))

                For i = 1 To 4
                    Set mRng = rngArray(i)
                    mRng.Merge
                    With mRng
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 90
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = True
                    End With

                Next i

                myRow = cRow + 2
                Exit For
            End If
        Next cRow
    End If
Next myRow

Application.DisplayAlerts = True

End Sub

```

**what I get is:**

![enter image description here](https://i.stack.imgur.com/uPaCb.png)

**_Question:_** **how to achieve this?**

Actually in my original data, the first three columns have data every 88 rows starting from row 3 and the column D should get merged every four rows.

----

Your code does not distinguish between the different columns in any way. If you know how many rows to merge you can simply search for cells and then do the merge based on column number. Here is one such approach which uses a pair of arrays to track how many rows to merge and then what formatting to apply.

You will need to change the row counts in the array definition. Sounds like you want (87,87,87,3) based on your edit. I did (11,11,11,3) to match your example though. This is the real fix to your code; it uses the `Column` number to determine how many rows to merge.

I also just typed some values into the spreadsheet and used `SpecialCells` to get only the cells with values. If your data matches your example, this works fine.

**Edit** includes unmerging cells first per OP request.

```
Sub MergeAllBasedOnColumn()

    Dim rng_cell As Range
    Dim arr_rows As Variant
    Dim arr_vert_format As Variant

    'change these to the actual number of rows
    'one number for each column A, B, C, D
    arr_rows = Array(11, 11, 11, 3)

    'change these if the formatting is different than example
    arr_vert_format = Array(True, True, True, False)

    'unmerge previously merged cells
    Cells.UnMerge

    'get the range of all cells, mine are all values
    For Each rng_cell In Range("A:D").SpecialCells(xlCellTypeConstants)

        'ignore the header row
        If rng_cell.Row > 2 Then

            'use column to get offset count
            Dim rng_merge As Range
            Set rng_merge = Range(rng_cell, rng_cell.Offset(arr_rows(rng_cell.Column - 1)))

            'merge cells
            rng_merge.Merge

            'apply formatting
            If arr_vert_format(rng_cell.Column - 1) Then
                'format for the rotated text (columns A:C)
                With rng_merge
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 90
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                End With
            Else
                'format for the other cells (column D)
                With rng_merge
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                End With
            End If
        End If
    Next rng_cell
End Sub

```

**Before**

![before](https://i.stack.imgur.com/4SAgG.png)

**After**

![after](https://i.stack.imgur.com/YZykf.png)
