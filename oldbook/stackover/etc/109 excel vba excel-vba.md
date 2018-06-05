# SO item 109
I have found the following code that simulates the MS Word Mail Merge functionality, but exclusively in Excel cells.

It was intended to work for a single cell value for each row in a range to be sent to another cell before printing, so that several copies can be printed, each with a different name on it for rationalization, such as with printing time cards to distribute to each employee.

However, I have not been able to figure out how to apply this to sending a range of three cells (in columns **`A:C`** from Employees table) for each row to another range of three cells (**`X50:X52`** in **`Sheet1`**, instead of just the one cell. Any ideas?

```
Sub Macro1()
   Dim lastRow As Integer '
   Dim r As Integer
   lastRow = Sheets("Employees").Cells(Rows.Count, "A").End(xlUp).Row

   For r = 1 To lastRow
   Sheets("Sheet1").Range("D1").Value = Sheets("Employees").Range("A" & r).Value
   ActiveWindow.SelectedSheets.PrintOutNext r
End Sub

```

----

You can use `Application.Transpose` to flip the row of data into a column of data. Below I also use `Resize` to get the 3 cells from columns `A:C`. `Resize` is much easier than try to build a range with `&` and column letters. It returns a new `Range` that is 3 columns large instead of the 1 column from before.

```
Sub Macro1()
   Dim lastRow As Integer '
   Dim r As Integer
   lastRow = Sheets("Employees").Cells(Rows.Count, "A").End(xlUp).Row

   For r = 1 To lastRow
       Sheets("Sheet1").Range("X50:X52").Value = _
            Application.Transpose(Sheets("Employees").Range("A" & r).Resize(, 3).Value)

       ActiveWindow.SelectedSheets.PrintOutNext r
   Next r
End Sub

```

Note that when selecting multiple cells with `Range` and calling `.Value`, you will get an array of values. In this case, the array is 1 row by 3 columns. `Application.Transpose` converts this data to be 3 rows by 1 column.
