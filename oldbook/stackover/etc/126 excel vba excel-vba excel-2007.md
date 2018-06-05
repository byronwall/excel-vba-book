# SO item 126
I have an excel workbook with 5 worksheets in each of this worksheets there is a table with values. I then delete some of the rows (using VBA) depending on user selection (using dropdown list). After my VBA code deleted all the unnecessary rows Excel states that I have "inconsistent column formulas" which I'd like to resolve with VBA before the user sees it.

Is there any way to do this with VBA?

I've searched Google the whole day now and still found nothing usefull and the only thing I'd have in my mind would be iterating through all Rows and Columns with formulas in it, checking, if the formula contains an error, which would definitely be super slow...

Note: If this counts as duplicate of [Find inconsistent formulas in Excel through VBA](http://stackoverflow.com/questions/24511585/find-inconsistent-formulas-in-excel-through-vba) I'm sorry, but the only answer there doesn't work with tables as data range

----

If you are trying to reset a formula in a `Table`, you can use the `DataBodyRange.Formula` property to reset the formula for the entire column. This addresses one way to get the `Inconsistent Calculated Column Formula`.

Example of the error was obtained by setting the formula for the column, changing one cell, and then telling Excel not to change the formula column after that edit.

[![enter image description here](https://i.stack.imgur.com/cueVa.png)](https://i.stack.imgur.com/cueVa.png)

To revert this back to a column formula (and remove the error), you can run VBA that changes the formula for the `DataBodyRange`.

**Code to change back**

```
Sub ResetTableColumnFormula()
    Dim listObj As ListObject
    For Each listObj In ActiveSheet.ListObjects
        listObj.ListColumns("b").DataBodyRange.Formula = "=[@a]"
    Next
End Sub

```

Note that I am being a bit lazy by iterating through all `ListObjects` instead of using the name of the table to get a reference. This works since there is only a single `Table` on the `Worksheet`.

**Formulas after that code runs:**

[![enter image description here](https://i.stack.imgur.com/UmzVO.png)](https://i.stack.imgur.com/UmzVO.png)

Note that this answer is very similar to the [answer here.](http://stackoverflow.com/a/13760891/4288101)
