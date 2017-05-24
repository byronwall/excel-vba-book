# SO item 037
**_Required:_** I'm trying to create a macro that filters **cell I22** for all _zero_ values, selects all those rows, deletes them and then unfilters again.

**_What I have:_** Currently I'm doing this in two different steps, since doing this in one step takes a couple hours (as it deletes row per row)

**Code (1)**: Autofilters to 'zero' and 'N/A', selects all of them and clears all the content. Next it clears the filter and sorts from largest to smallest. This way excel doesn't have to delete each row separately making the process faster.

**Code (2)**: Deletes all the blank rows.

I have the impression this code is not exactly efficient and too long given the task it needs to do. Is it possible to combine these into one code?

**_Code (1)_**

```
Sub clearalldemandzero()

clearalldemandzero Macro

ActiveWindow.SmallScroll Down:=15
Range("A26:EU26").Select
Selection.AutoFilter
ActiveWindow.SmallScroll ToRight:=3
ActiveSheet.Range("$A$26:$EU$5999").AutoFilter Field:=9, Criteria1:="=0.00" _
    , Operator:=xlOr, Criteria2:="=#N/A"
Rows("27:27").Select
Range("D27").Activate
Range(Selection, Selection.End(xlDown)).Select
Selection.Clear
ActiveSheet.ShowAllData
Range("H28").Select
ActiveWorkbook.Worksheets("Solver 4").AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Solver 4").AutoFilter.Sort.SortFields.Add Key:= _
    Range("I26:I5999"), SortOn:=xlSortOnValues, Order:=xlDescending, _
    DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Solver 4").AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
End Sub

```

**_Code (2)_**

```
Sub DeleteBlankRows3()

'Deletes the entire row within the selection if the ENTIRE row contains no data.'

Dim Rw As Range

If WorksheetFunction.CountA(Selection) = 0 Then

    MsgBox "No data found", vbOKOnly, "OzGrid.com"
    Exit Sub
End If

With Application

    .Calculation = xlCalculationManual
    .ScreenUpdating = False
    Selection.SpecialCells(xlCellTypeBlanks).Select

    For Each Rw In Selection.Rows

        If WorksheetFunction.CountA(Selection.EntireRow) = 0 Then

            Selection.EntireRow.Delete
        End If
    Next Rw

    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True

End With

End Sub

```

----

If your code to Select the filtered data is working, you can simply delete all the rows at that step in one shot. The key is to use `SpecialCells` and only select visible cells. Then you can get the `EntireRow` and `Delete` it.

The relevant line of code to add would be this:

```
Selection.SpecialCells(xlCellTypeVisible).EntireRow.Delete

```

The modification to code 1 in its entirety should be:

```
Sub clearalldemandzero()

    clearalldemandzero Macro

    ActiveWindow.SmallScroll Down:=15
    Range("A26:EU26").Select
    Selection.AutoFilter
    ActiveWindow.SmallScroll ToRight:=3
    ActiveSheet.Range("$A$26:$EU$5999").AutoFilter Field:=9, Criteria1:="=0.00" _
        , Operator:=xlOr, Criteria2:="=#N/A"
    Rows("27:27").Select
    Range("D27").Activate
    Range(Selection, Selection.End(xlDown)).Select

    Selection.SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

End Sub

```

As a side note, you should generally work to avoid using `Select` `Selection` and other things that interface with the Excel UI like that. I did not try to fix those issues here since it seems that your code is generally working. Reference to that issue: [How to avoid using Select in Excel VBA macros](http://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba-macros)
