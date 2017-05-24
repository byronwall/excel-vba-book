# SO item 015
I'm new here and I apologize in advance in my question isn't clear... I couldn't find the answer after some research...

I'm looking for a way to go through all the cells of column "R" and if one cell on a given row contains "Y", then the values of cells at columns "W","X" and "Y" will take the same value as the columns "F","G" and "H" (always at the same row).

The goal is to have a button that will execute the VBA code in order to do this (instead of having to copy/paste all the time).

Thank you very much in advance for your help.

A poor ignorant but motivated VBA beginner...

----

Here is VBA which will do what you want. It takes advantage of the replacement operation being cells that are next to each other by using `Resize`.

Highlights

1.  Iterates through each cell in column R. I used Intersect with the UsedRange on the sheet so that it only goes through cells that have values in them (instead of all the way to the end).
2.  Checks for "Y" using `InStr`.
3.  Replaces the contents of columns WXY with values from columns FGH. Since they are contiguous, I did it all in one step with `Resize`.

Code:

```
Sub ReplaceValuesBasedOnColumn()

    Dim rng_search As Range
    Dim rng_cell As Range

    'start on column R, assume correct sheet is open
    Set rng_search = Range("R:R")
    For Each rng_cell In Intersect(rng_search, rng_search.Parent.UsedRange)

        'search for a Y, case sensitive
        If InStr(rng_cell, "Y") > 0 Then

            'update the columns as desired
            'takes advantage of cells being next to each other
            Range("W" & rng_cell.Row).Resize(1, 3).Value = Range("F" & rng_cell.Row).Resize(1, 3).Value

        End If
    Next rng_cell
End Sub

```

I tested it on my end, and it works, producing the following after running:

![picture after running](https://i.stack.imgur.com/9sII5.png)
