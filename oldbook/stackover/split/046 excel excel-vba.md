# SO item 046
I am looking to insert a new row below a frozen header row at the top of a spreadsheet. The issue I face is the amount of rows in the header is ever changing but I always want the row to be inserted at the first line below the header. Is there a flag in the row that says its frozen? which I could just count the amount of rows with said flag, add 1 and insert row. Any help would be very helpful.

Matt

----

If you are using `FreezePanes` then I think you go this route:

```
Sub InsertRowBelowHeader()
    Rows(ActiveWindow.Panes(1).VisibleRange.Rows.Count + 1).Insert
End Sub

```

**Before**, the freeze line is below row 5\. Freeze pane was done on cell `A6`

![before](https://i.stack.imgur.com/EXxls.png)

**After**, a row is added to split a/b

![after](https://i.stack.imgur.com/IrT1V.png)

Here is a relevant discussion which came up on Google for freeze panes and VBA. [http://www.mrexcel.com/forum/excel-questions/275645-identifying-freeze-panes-position-sheet-using-visual-basic-applications.html](http://www.mrexcel.com/forum/excel-questions/275645-identifying-freeze-panes-position-sheet-using-visual-basic-applications.html)
