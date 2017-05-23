# SO item 002
I am trying to use a macro to delete data that follows after a certain word/phrase in excel.

The problem is the cell position can vary depending on how many lines there are in the spreadsheet after the report has been run/exported. I need a solution, if possible, to target a certain phrase (using the find function i presume) and delete 30 cells down from that wherever the text may be (no cet cell).

Is this possible?

----

This is possible, but the specifics of what you want are not the clearest. Here is some code that will get you started. It looks for a given phrase in all of the cells on the sheet. If it finds it, it adds the 30 cells below that one into a group that is slated for deletion. Once it searches all the cells, it deletes the groups that were collected.

Since we are iterating through the collection that we want to delete from, I collect the cells to delete (using Union) and delete them after the loop closes.

This may not be the fastest code since it searches all cells in the UsedRange, but it should work for most uses.

```
Sub deleteBelowPhrase()

    Dim cell As Range
    Dim sht As Worksheet

    Dim phrase As String
    Dim deleteRows As Integer

    'these are the parameters
    phrase = "DELETE BELOW ME"
    deleteRows = 30

    'assumes you are searching the active sheet
    Set sht = ActiveSheet

    Dim del_range As Range
    Dim del_cell As Range

    For Each cell In sht.UsedRange

        'check for the phrase in the cell
        If InStr(cell.Value2, phrase) > 0 Then

            'get a reference to the range to delete
            Set del_cell = sht.Range(cell.Offset(1), cell.Offset(deleteRows))

            'create object to delete or add to existing
            If del_range Is Nothing Then
                Set del_range = del_cell
            Else
                Set del_range = Union(del_range, del_cell)
            End If
        End If

    Next cell

    'delete the block of cells that are collected
    del_range.Delete xlShiftUp

End Sub

```

# Before

![before](https://i.stack.imgur.com/Qfuxv.png)

# After

![after](https://i.stack.imgur.com/hHYMs.png)
