# SO item 088
Excel 2013\. I am using 3 combo boxes to change filters on the pivot table. My first combo box has "Project1", "Project2" & All. My second combo box has "Customer1", "Customer2" & All. My third combo box has "Country1", "Country2" & All.

I am using 9 pivot tables, all of them have filters as [Project], [Customer], [Country].

My intention is to change first combo box to Project1 & all the pivot tables filter should change as Project1.I am successfully able to do that.

However when I select the first combo box as "All". First Combo box cell link to Y1\. I get VBA Run time error 1004: Application-defined or object-defined error.

```
Sub ProjectName()

ActiveSheet.PivotTables("PVT1").PivotFields("Project Name").ClearAllFilters
ActiveSheet.PivotTables("PVT2").PivotFields("Project Name").ClearAllFilters
ActiveSheet.PivotTables("PVT3").PivotFields("Project Name").ClearAllFilters

    ActiveSheet.PivotTables("PVT1").PivotFields("Project Name").CurrentPage = Range("Y1").Text
    ActiveSheet.PivotTables("PVT2").PivotFields("Project Name").CurrentPage = Range("Y1").Text
    ActiveSheet.PivotTables("PVT3").PivotFields("Project Name").CurrentPage = Range("Y1").Text 

```

----

Since the first three lines of code go without issue, I will assume that the Pivot Table `PVT1` and the Field `Project Name` all exist. This places the error somewhere after that.

For the call to `.CurrentPage` you will get a 1004 error for the following reasons:

*   Using this to try and filter any field that is not set as a `Report Filter`. You cannot use the `CurrentPage` to filter any rows or columns
*   Setting a value which does not exist in the list of possible values

On the second point, this is where the call to `Range` might be relevant.

*   Verify that the value there exists in the list of possible ones.
*   Also be aware you are using `.Text` which will use the _display_ value of the cell and not its underlying `.Value`

To resolve these issues, there are a couple of options:

*   For the case where you want to filter data that is on the row or column (and not in the filters section) you can go through `PivotItems` and set `Visible = True/False`
*   You can also set a label filter from VBA if you want that instead of the manual filter
*   If you want to check for a value existing in the `CurrentPage`, you can iterate the `PivotItems` for that `PivotField` and check that one matches. The code for that is very similar to the `For Each` loop with the check on value, just don't set `Visible`.

**Code** for setting a filter on a row or column

```
Sub FilterPivotField()

    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables("PVT1")

    Dim pf As PivotField
    Set pf = pt.PivotFields("C")

    pf.ClearAllFilters

    'slow iterates all items and sets Visible (manual filter)
    Dim pi As PivotItem
    For Each pi In pf.PivotItems
        pi.Visible = (pi.Name = Range("J2"))
    Next

    'fast way sets a label filter
    pf.PivotFilters.Add2 Type:=xlCaptionEquals, Value1:=Range("J2")

End Sub

```

**Picture of ranges**

![enter image description here](https://i.stack.imgur.com/U2iYX.png)
