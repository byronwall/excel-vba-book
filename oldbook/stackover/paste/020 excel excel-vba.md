# SO item 020
I have a data table which is sorted on descending order in column F. I then need to copy the top 5 rows, but only data from Column A, B, D, and F (not the headers). See pictures.

```
Sub top5()

Sheets("Sheet1").Select

If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
ActiveSheet.ShowAllData
End If

ActiveSheet.Range("$A$4:$T$321").AutoFilter Field:=3, Criteria1:="Dave"
ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields. _
    Clear
ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add _
    Key:=Range("F4:F321"), SortOn:=xlSortOnValues, Order:=xlDescending, _
    DataOption:=xlSortTextAsNumbers
With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

' This copy-paste part does what its supposed to, but only for the specific 
' cells.  Its not generalised and I will have to repeat this operation
' several times for different people
Sheets("Sheet1").Select
Range("A3:B15").Select
Selection.Copy

Sheets("Sheet2").Select
Range("A3").Select
ActiveSheet.Paste

Sheets("Sheet1").Select
Range("D3:D15").Select
Application.CutCopyMode = False
Selection.Copy

Sheets("Sheet2").Select
Range("C3").Select
ActiveSheet.Paste

Sheets("Sheet1").Select
Range("F3:F15").Select
Application.CutCopyMode = False
Selection.Copy

Sheets("Sheet2").Select
Range("D3").Select
ActiveSheet.Paste
Application.CutCopyMode = False

End Sub

```

I thought about trying to adapt this snippet of code below using visible cells function, but I'm stuck and I can't find anything on the net which fits.

```
' This selects all rows (plus 1, probably due to offset), I only want parts of from the top 5.
Sheets("Sheet1").Select
ActiveSheet.Range("$A$4:$B$321").Offset(1, 0).SpecialCells(xlCellTypeVisible).Select
Selection.Copy
Sheets("Sheet2").Select
Range("A3").Select
ActiveSheet.Paste

Sheets("Sheet1").Select
ActiveSheet.Range("$D$4:$D$321").Offset(1, 0).SpecialCells(xlCellTypeVisible).Select
Selection.Copy
Sheets("Sheet2").Select
Range("C3").Select
ActiveSheet.Paste

```

I hope my example makes sense and I really appreciate your help!

![Sample Excel table](https://i.stack.imgur.com/pitTN.png)

Note: The heading names are only the same in the two tables to show that the data is the same. The headers are NOT supposed to be copied. In addition, there is an extra column/white space in the second table. A solution should include this.

![Data copied to new table](https://i.stack.imgur.com/XXCK4.png)

----

A quick way to do this is to use `Union` and `Intersect` to only copy the cells that you want. If you are pasting values (or the data is not a formula to start), this works well. Thinking about it, it builds a range of columns to keep using `Union` and then `Intersect` that with the first 5 rows of data with 2 header rows. The result is a copy of only the data you want with formatting intact.

**Edit only process visible rows, grabbing the header, and then the first 5 below the header rows**

```
Sub CopyTopFiveFromSpecificColumns()

    'set up the headers first to keep
    Dim rng_top5 As Range
    Set rng_top5 = Range("3:4").EntireRow

    Dim int_index As Integer
    'start below the headers and keep all the visible cells
    For Each cell In Intersect( _
        ActiveSheet.UsedRange.Offset(5), _
        Range("A:A").SpecialCells(xlCellTypeVisible))

        'add row to keepers
        Set rng_top5 = Union(rng_top5, cell.EntireRow)

        'track how many items have been stored
        int_index = int_index + 1
        If int_index >= 5 Then
            Exit For
        End If
    Next cell

    'copy only certain columns of the keepers
    Intersect(rng_top5, _
        Union(Range("A:A"), _
                Range("B:B"), _
                Range("D:D"), _
                Range("F:F"))).Copy

    'using Sheet2 here, you can set to wherever, works if data is not formulas
    Range("Sheet2!A1").PasteSpecial xlPasteAll

    'if the data contains formulas, use this route
    'Range("Sheet2!A1").PasteSpecial xlPasteValues
    'Range("Sheet2!A1").PasteSpecial xlPasteFormats

End Sub

```

Here is the result I get from some dummy data set up in the same ranges as the picture above.

**Sheet1 with copied range visible**

![Sheet1](https://i.stack.imgur.com/6KLmb.png)

**Sheet2 with pasted data**

![Sheet2](https://i.stack.imgur.com/ZfjiT.png)
