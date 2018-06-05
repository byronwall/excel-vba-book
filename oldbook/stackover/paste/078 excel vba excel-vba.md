# SO item 078
I am trying to copy a set of filtered data from one sheet to the bottom of another sheet. And my code works great except for the first time upon opening the file I get a:

> Run Time error 1004

If I quit the debugger and re-run the macro it works great.
Here is my code: noted where the problem occurs.

```
Sub MoveData_Click()
    'Select the filtered alarm data and paste on the master spreadsheet
    Sheets("DailyGen").Select
    ActiveSheet.UsedRange.Offset(5, 0).SpecialCells _
        (xlCellTypeVisible).Copy

    Sheets("2015 Master").Select

    If ActiveWorkbook.ActiveSheet.FilterMode _
    Or ActiveWorkbook.ActiveSheet.AutoFilterMode Then
        ActiveWorkbook.ActiveSheet.ShowAllData
    End If

    Range("C4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, -2).Range("A1").Select
    ActiveSheet.Paste '~~> THIS IS WHERE IT ERRORS

    'Sort newest to oldest in the date column

    ActiveWorkbook.Worksheets("2015 Master").ListObjects("Table44").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("2015 Master").ListObjects("Table44").Sort.SortFields.Add _
        Key:=Range("Table44[[#All],[Active Time]]"), _
        SortOn:=xlSortOnValues, 
        Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("2015 Master").ListObjects("Table44").Sort
       .Header = xlYes
       .MatchCase = False
       .Orientation = xlTopToBottom
       .SortMethod = xlPinYin
       .Apply
    End With
End Sub

```

----

When you `ShowAllData` (same as `Data->Clear` in the Filter section) you are emptying the clipboard and telling Excel to forget about the copied `Range`. Do it outside of VBA to confirm if you want. Excel loves to empty the clipboard when you edit a cell or do much of anything other than selecting.

To fix, do the `Copy` after the `ShowAllData`. In your case, you will have to `Select` the `Worksheet` back and forth.

You should generally work to avoid using `Select` and `Activate` for your VBA. [See this post for details.](http://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba-macros)

Here is the final code with the changes made:

```
Sub MoveData_Click()

'Select the filtered alarm data and paste on the master spreadsheet

Sheets("2015 Master").Select
If ActiveWorkbook.ActiveSheet.FilterMode Or ActiveWorkbook.ActiveSheet.AutoFilterMode Then
ActiveWorkbook.ActiveSheet.ShowAllData
End If

Sheets("DailyGen").Select
ActiveSheet.UsedRange.Offset(5, 0).SpecialCells _
    (xlCellTypeVisible).Copy

Sheets("2015 Master").Select
Range("C4").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, -2).Range("A1").Select
ActiveSheet.Paste

 'Sort newest to oldest in the date column

ActiveWorkbook.Worksheets("2015 Master").ListObjects("Table44").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("2015 Master").ListObjects("Table44").Sort.SortFields.Add _
    Key:=Range("Table44[[#All],[Active Time]]"), SortOn:=xlSortOnValues, Order _
     :=xlDescending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("2015 Master").ListObjects("Table44").Sort
   .Header = xlYes
   .MatchCase = False
   .Orientation = xlTopToBottom
   .SortMethod = xlPinYin
   .Apply
End With

End Sub

```
