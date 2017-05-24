# SO item 007
I have some hierarhical dictionary like this

![enter image description here](https://i.stack.imgur.com/JRdw0.jpg)

And unloaded this dictionary into Excel file like this

![enter image description here](https://i.stack.imgur.com/DL7Cy.jpg)

Now I have to group Excel rows in hierarhical manner.

I know that Excel has range().group function. But is there any way to do so quick and easy or I have to scan all excel file and select rows to group in manual way?

----

I tested this out on some dummy examples and it worked well. The idea is to loop through the cells in your range of interest and group them if the level is greater in rows below the current cell. It works its way down which builds the hierarchy. You can then use the grouping buttons to show different levels.

This code works if the level column has no breaks (since I am using Range.End to get the last cell).

```
Sub GroupBasedOnLevel()

    Dim rng_cells As Range
    Dim rng_start As Range
    Dim rng_end As Range

    'set up some ranges, change rng_start to be appropriate
    Set rng_start = Range("A2")
    Set rng_end = rng_start.End(xlDown)
    Set rng_cells = Range(rng_start, rng_end)

    'clear previous outline
    Cells.ClearOutline

    'loop through level cells and group based on values below
    Dim cell As Range
    For Each cell In rng_cells

        'get value of cell and start checking below it
        Dim row_off As Integer
        row_off = 1

        'loop ensures level is greater below and cells are within range
        Do While cell.Offset(row_off) > cell And cell.Offset(row_off).Row <= rng_end.Row
            row_off = row_off + 1
        Loop

        'do the grouping if there are more than 1 cells below
        If row_off > 1 Then
            Range(cell.Offset(1), cell.Offset(row_off - 1)).EntireRow.Group
        End If
    Next cell
End Sub

```

# Results

## starting point

![starting point](https://i.stack.imgur.com/lBwAZ.png)

## full grouping

![full grouping](https://i.stack.imgur.com/39o9K.png)

## collapsed to level 1

![collapsed group](https://i.stack.imgur.com/uxHsZ.png)
