# SO item 093
I'm trying to create a PivotTable in which a double click on a value leads the user to the filtered source sheet with the rows that this value represents, rather than a new sheet with the underlying data.

This is how far I've gotten, but I'm having issues extracting the relevant row and column names / values, as well as the filters currently active in the pivottable.

```
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim rng As Range
    Dim wks As Worksheet
    Dim pt As PivotTable

    ' Based on http://stackoverflow.com/questions/12526638/how-can-you-control-what-happens-when-you-double-click-a-pivot-table-entry-in-ex
    Set wks = Target.Worksheet
    For Each pt In wks.PivotTables()
        Set rng = Range(pt.TableRange1.Address)
        If Not Intersect(Target, rng) Is Nothing Then
            Cancel = True
        End If
    Next  

     ' Source: http://www.mrexcel.com/forum/excel-questions/778468-modify-pivottable-double-click-behavior.html
     On Error GoTo ExitNow
     With Target.PivotCell
         If .PivotCellType = xlPivotCellValue And _
             .PivotTable.PivotCache.SourceType = xlDatabase Then
                 SourceTable = .PivotTable.SourceData
                 MsgBox SourceTable
                 ' I found the sourcetable, how would I collect the row/column
                 ' names and values in order to filter this table?
         End If
     End With

ExitNow: Exit Sub
End Sub

```

In order to filter the source sheet, I need to extract the following characteristics upon a double click:

*   The filters active in the current PivotTable (the original** 'Fieldname' and the relevant filters)
*   The original** headers and row names and values relevant to the aggregate being selected (e.g. FieldX = 2013, FieldY="X"), that will enable me to filter the source sheet and present the underlying rows.

** Note that I'm not sure if this is relevant, but I extensively stumble upon PivotTables in which the row names shown are not the same as those in the source datasheet (by manually renaming them in the PivotTable). Also, is it possible to extract the 'groupings' created in the PivotTables?

Using these characteristics, the VBA for locating the source data and applying the relevant filters should be relatively straightforward. In most cases, the source table is an 'Excel Table', if this is relevant.

Any help is greatly appreciated.

----

The solution to this depends greatly on the filters you have in place. The way that `PivotFilters` are defined is different from the way that AutoFilters are defined. This means that you will need to do a translation for each type of filter that is in place.

AutoFilters do all of their magic in the `Criteria1` whereas the `PivotFilters` have a `FilterType` and `Value1` to make it work. This is the translation step.

For simple equality, it is fairly easy and that is the code included below. It address the issue of how to find the column header and set the filter.

```
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    Dim pt As PivotTable

    Dim wks As Worksheet
    Set wks = Target.Worksheet
    For Each pt In wks.PivotTables()
        If Not Intersect(Target, pt.TableRange1) Is Nothing Then
            Cancel = True
        End If
    Next

    If Cancel <> True Then
        Exit Sub
    End If

    Set pt = Target.PivotCell.PivotTable

    Dim rng As Range
    Set rng = Application.Range(Application.ConvertFormula(pt.SourceData, xlR1C1, xlA1))

    Dim sht_rng As Worksheet
    Set sht_rng = rng.Parent
    sht_rng.AutoFilterMode = False

    Dim pf As PivotField
    For Each pf In pt.PivotFields
        Dim pfil As PivotFilter
        For Each pfil In pf.PivotFilters
            If pfil.FilterType = xlCaptionEquals Then
                rng.AutoFilter Field:=Application.Match(pf.SourceName, rng.Rows(1), 0), Criteria1:=pfil.Value1
            End If
        Next pfil
    Next pf

    sht_rng.Activate
    rng.Cells(1, 1).Select
End Sub

```

Couple of notes:

*   I am using `PivotTable.SourceData` to get the range of cells that are involved. This returns a value in R1C1 notation, so I convert it to A1 notation using `Application.ConvertFormula`. I then need to use `Application.Range` to look up this string. (Since this code is executing within the scope of a specific `Worksheet` you need to add `Application` here so it expands the scope of the search)
*   After that it is a simple matter of iterating through all the PivotFields and their PivotFilters.
*   Inside that loop, then you need to find the column header (using `Application.Match` in the header row: `.Rows(1)`) and add the filter. This is where the conversion steps are required. You could do a `Select... Case` for each supported type of filter.
*   You might also want to check out `CurrentPage` if any of the fields is a filter instead of a row/column.
*   Finally, it is possible for there to be manual filters instead of the label filters which I am iterating through. You can loop through `PivotItems` and check for `Visible` if you want those.

Hopefully this code gets you started but also hints at the complexity of the task involved. You will likely want to limit yourself to supporting specific types of filters.

**Pictures of Pivot and data**

_pivot table with filters_

![pivot table with filters](https://i.stack.imgur.com/dFOjH.png)

_all data_

![all data](https://i.stack.imgur.com/co2KP.png)

_filtered data_

![filtered data](https://i.stack.imgur.com/O9FFE.png)
