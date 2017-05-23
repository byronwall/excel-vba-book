# SO item 115
I have a spreadsheet with many numbers, and I want the cells with the same numbers to be moved to the same row. Currently, my spreadsheet looks something like this:

```
*  May     Jun     Jul     Aug     Sep     Oct
* 10584   10589   10584   10584   10589   10589
* 10589   11202   10589   10589   11202   11202
* 11202   9799    11202   11202   11677   11677
*                         11677     

```

I would like to have some vba code to organize the data so that cells with the same value are on the same row, so it should look like this:

```
*  May     Jun     Jul     Aug     Sep     Oct
*         9799
* 10584           10584   10584  
* 10589   10589   10589   10589   10589   10589
* 11202   11202   11202   11202   11202   11202
*                         11677   11677   11677

```

With empty cells in the places with no numbers. I tried searching through the forum but I wasn't able to find anything similiar enough. I would really apreciate any help on this. Thanks for your time.

----

Here is an approach that works on a block of data of arbitrary size. It works by sorting the columns and then shifting the cells down if they are not equal to the smallest value in the row.

The only real parameter here to adjust is the starting cell: `rng_start` which is initially set to the `ActiveCell`. This code also uses `CurrentRegion` so the data needs to be a block... or you can redefine those couple of lines.

**Code**

```
Sub SortAndPutSameValuesInSameRow()

    'get data ranges
    Dim rng_start As Range
    Set rng_start = ActiveCell

    Dim rng_data As Range
    Set rng_data = rng_start.CurrentRegion
    Set rng_data = Intersect(rng_data, rng_data.Offset(1))

    'sort by column
    Dim rng_col As Range
    For Each rng_col In rng_data.Columns
        rng_col.Sort Key1:=rng_col
    Next

    'iterate through rows and arrange
    Dim rng_row As Range
    For Each rng_row In rng_data.Rows
        Dim rng_cell As Range
        For Each rng_cell In rng_row.Cells
            If rng_cell.Value <> Application.WorksheetFunction.min(rng_row) Then
                rng_cell.Insert xlShiftDown
            End If
        Next

        'break out if cell goes past data
        If Intersect(rng_row, rng_start.CurrentRegion) Is Nothing Then
            Exit For
        End If
    Next
End Sub

```

**How it works**

The main idea here is that once the columns are sorted, you just need to move values down so that only the smallest value is kept in each row. This logic also ensures that all of the same values are in the same row. Note that if there are duplicate values, you will get a row of matching values and then duplicate values (which would also match if repeated in multiple columns). Specific comments:

*   The top half of the code is setting things up for the iteration section below. It grabs the block of data and builds a range that excludes the headers.
*   With the block of data, it first goes through each column and sorts them in turn.
*   Once sorted, it goes through each row of the data and checks if the current value is equal to the minimum value in the row.
*   If so, then that cell can stay put. If not, the values need to shift down to make a blank cell.
*   Finally, there is a check to bust out of the loop when needed. This is a little odd in a `For Each` loop but is required because the size of the range is changing as it iterates (because of `Insert`).

Since I am using `Rows` and `Columns`, this code will work for data anywhere on the sheet and for as many columns as you want.

**Pictures of before/after** show results with your data

_before_

![before](https://i.stack.imgur.com/AxJ6j.png)

_after_

![enter image description here](https://i.stack.imgur.com/LSM0W.png)
