# SO item 070
I'm trying to figure out how to split rows of data where columns B,C,D in the row contain multiple lines and others do not. I've figured out how to split the multi-line cells if I copy just those columns into a new sheet, manually insert rows, and then run the macro below (that's just for column A), but I'm lost at coding the rest.

Here's what the data looks like: ![enter image description here](https://i.stack.imgur.com/Y0Y5p.png)

So for row 2, I need it split into 6 rows (one for each line in cell B2) with the text in cell A2 in A2:A8\. I also need columns C and D split the same as B, and then **columns E:CP** the same as column A.

Here is the code I have for splitting the cells in columns B,C,D:

```
Dim iPtr As Integer
Dim iBreak As Integer
Dim myVar As Integer
Dim strTemp As String
Dim iRow As Integer
iRow = 0
For iPtr = 1 To Cells(Rows.Count, col).End(xlUp).Row
    strTemp = Cells(iPtr1, 1)
    iBreak = InStr(strTemp, vbLf)
    Range("C1").Value = iBreak
        Do Until iBreak = 0
        If Len(Trim(Left(strTemp, iBreak - 1))) > 0 Then
            iRow = iRow + 1
            Cells(iRow, 2) = Left(strTemp, iBreak - 1)
        End If
        strTemp = Mid(strTemp, iBreak + 1)
        iBreak = InStr(strTemp, vbLf)
    Loop
    If Len(Trim(strTemp)) > 0 Then
        iRow = iRow + 1
        Cells(iRow, 2) = strTemp
    End If
Next iPtr
End Sub

```

Here is a link to an example file (note this file has 4 rows, the actual sheet has over 600): [https://www.dropbox.com/s/46j9ks9q43gwzo4/Example%20Data.xlsx?dl=0](https://www.dropbox.com/s/46j9ks9q43gwzo4/Example%20Data.xlsx?dl=0)

----

This is a fairly interesting question and something I have seen variations of before. I went ahead and wrote up a general solution for it since it seems like a useful bit of code to keep for myself.

There are pretty much only two assumptions I make about the data:

*   Returns are represented by `Chr(10)` or which is the `vbLf` constant.
*   Data that belongs with a lower row has enough returns in it to make it line up. This appears to be your case since there are return characters which appear to make things line up like you want.

**Pictures of the output**, zoomed out to show all the data for `A:D`. Note that the code below **processes all of the columns by default and outputs to a new sheet**. You can limit the columns if you want, but it was _too_ tempting to make it general.

![output of the code](https://i.stack.imgur.com/6bo3f.png)

**Code**

```
Sub SplitByRowsAndFillBlanks()

    'process the whole sheet, could be
    'Intersect(Range("B:D"), ActiveSheet.UsedRange)
    'if you just want those columns
    Dim rng_all_data As Range
    Set rng_all_data = Range("A1").CurrentRegion

    Dim int_row As Integer
    int_row = 0

    'create new sheet for output
    Dim sht_out As Worksheet
    Set sht_out = Worksheets.Add

    Dim rng_row As Range
    For Each rng_row In rng_all_data.Rows

        Dim int_col As Integer
        int_col = 0

        Dim int_max_splits As Integer
        int_max_splits = 0

        Dim rng_col As Range
        For Each rng_col In rng_row.Columns

            'splits for current column
            Dim col_parts As Variant
            col_parts = Split(rng_col, vbLf)

            'check if new max row count
            If UBound(col_parts) > int_max_splits Then
                int_max_splits = UBound(col_parts)
            End If

            'fill the data into the new sheet, tranpose row array to columns
            sht_out.Range("A1").Offset(int_row, int_col).Resize(UBound(col_parts) + 1) = Application.Transpose(col_parts)

            int_col = int_col + 1
        Next

        'max sure new rows added for total length
        int_row = int_row + int_max_splits + 1
    Next

    'go through all blank cells and fill with value from above
    Dim rng_blank As Range
    For Each rng_blank In sht_out.Cells.SpecialCells(xlCellTypeBlanks)
        rng_blank = rng_blank.End(xlUp)
    Next

End Sub

```

**How it works**

There are comments within the code to highlight what is going on. Here is a high level overview:

*   Overall, we iterate through each row of the data, processing all of the columns individually.
*   The text of the current cell is `Split` using the `vbLf`. This gives an array of all the individual lines.
*   A counter is tracking the maximum number of rows that were added (really this is `rows-1` since these arrays are `0-indexed`.
*   Now the data can be output to the new sheet. This is easy because we can just dump the array that `Split` created for us. The only tricky part is getting it to the right spot on the sheet. To that end, there is a counter for the current column offset and a global counter to determine how many total rows need to be offset. The `Offset` moves us to the right cell; the `Resize` ensures that all of the rows are output. Finally, `Application.Transpose` is needed because `Split` returns a row array and we're dumping a column.
*   Update the counters. Column offset is incremented every time. The row offset is updated to add enough rows to cover the last maximum (`+1` since this is `0-indexed`)
*   Finally, I get to [use my waterfall fill (your previous question)](http://stackoverflow.com/questions/30537813/count-lines-of-text-in-a-cell/30538117#30538117) on all of the blanks cells that were created to ensure no blanks. I forgo error checking because I assume blanks exist.
