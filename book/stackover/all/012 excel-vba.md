# SO item 012
I'm trying to copy masses of information from one spreadsheet to another to make it easier to print out on one piece of paper. All the data is set out in sequence and in columns and they need to be printed as such.

I'm trying to create a userform to speed this up by copying different column ranges and pasting them in to another spreadsheet in the exact same format but in columns of 50 cells and a maximum of 4 columns per sheet of paper.

This is what I've got so far, but it only copies the first cell:

```
Private Sub UserForm_Click()

    UserForm1.RefEdit1.Text = Selection.Address

End Sub
Private Sub CommandButton1_Click()

    Dim addr As String, rng
    Dim tgtWb As Workbook
    Dim tgtWs As Worksheet
    Dim icol As Long
    Dim irow As Long

    Set tgtWb = ThisWorkbook
    Set tgtWs = tgtWb.Sheets("Sheet1")

    addr = RefEdit1.Value
    Set rng = Range(addr)

    icol = tgtWs.Cells(Rows.Count, 1) _
    .End(xlUp).Offset(0, 0).Column

    tgtWs.Cells(1, icol).Value = rng.Value

End Sub

```

Any help would be greatly appreciated.

----

Your approach for outputting the data is only referencing a single cell. You use `.Cells(1,icol)` which will only reference a single cell (in row 1, and a single column).

In order to output the data to a larger range, you need to reference a larger range. The easiest way to do this is probably via `Resize()` using the size of the RefEdit range.

I believe this will work for you. I changed the last line to include a call to `Resize`.

```
Private Sub CommandButton1_Click()

    Dim addr As String, rng
    Dim tgtWb As Workbook
    Dim tgtWs As Worksheet
    Dim icol As Long
    Dim irow As Long

    Set tgtWb = ThisWorkbook
    Set tgtWs = tgtWb.Sheets("Sheet1")

    addr = RefEdit1.Value
    Set rng = Range(addr)

    icol = tgtWs.Cells(Rows.Count, 1) _
    .End(xlUp).Offset(0, 0).Column

    tgtWs.Cells(1, icol).Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

End Sub

```

**Edit:** I went ahead and created a dummy example to test this out:

![workbooks](https://i.stack.imgur.com/DuZcM.png)

Click the button and it pastes

![click](https://i.stack.imgur.com/lYRjh.png)
