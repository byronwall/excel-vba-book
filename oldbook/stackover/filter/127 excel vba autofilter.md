# SO item 127
I am trying to filter a range of values and based on my criteria, at times I might have no data that fits my criteria. In that case, I do not want to copy any data from the filtered data. If there is filtered data, then I would like to copy it.

Here is my code:

```
With Workbooks(KGRReport).Worksheets(spreadSheetName).Range("A1:I" & lastrowinSpreadSheet)
    .AutoFilter Field:=3, Criteria1:=LimitCriteria, Operator:=xlFilterValues 'Do the filtering for Limit
     .AutoFilter Field:=9, Criteria1:=UtilizationCriteria, Operator:=xlFilterValues 'Do the filtering for Bank/NonBank
End With

'Clear the template
 Workbooks(mainwb).Worksheets("Template").Activate
 Workbooks(mainwb).Worksheets("Template").Rows(7 & ":" & Rows.Count).Delete

 'Copy the filtered data
 Workbooks(KGRReport).Activate
 Set myRange = Workbooks(KGRReport).Worksheets(spreadSheetName).Range("B2:H" & lastrowinSpreadSheet).SpecialCells(xlVisible)
 For Each myArea In myRange.Areas
     For Each rw In myArea.Rows
           strFltrdRng = strFltrdRng & rw.Address & ","
     Next
 Next

 strFltrdRng = Left(strFltrdRng, Len(strFltrdRng) - 1)
 Set myFltrdRange = Range(strFltrdRng)
 myFltrdRange.Copy
 strFltrdRng = ""

```

It is giving me an error at

```
Set myRange = Workbooks(KGRReport).Worksheets(spreadSheetName).Range("B2:H" & lastrowinSpreadSheet).SpecialCells(xlVisible)

```

When there is no data at all, it is returning an error: "No cells found".

Tried error handling like this post: [1004 Error: No cells were found, easy solution?](http://stackoverflow.com/questions/25380886/1004-error-no-cells-were-found-easy-solution)

But it was not helping. Need some guidance on how to solve this.

----

**An approach without the error handling**

It is possible to build the `AutoFilter` in a way that does not throw the error if nothing is found. The trick is to **include the header row in the call to the `SpecialCells`**. This will ensure that at least 1 row is visible even if nothing is found (Excel will not hide the header row). This prevents the error from jamming up execution and gives you a set of cells to check if data was found.

To check if the resulting range has data, you need to check `Rows.Count > 1 Or Areas.Count > 1`. This handles the two possible cases where your data is found directly under the header or in a discontinuous range below the header row. Either result means that the `AutoFilter` found valid rows.

Once you check that data was found, you can then do the desired call to `SpecialCells` on the data only without concern for an error.

**Sample data [column C (field 2) will be filtered]:**

[![random data](https://i.stack.imgur.com/iKHtN.png)](https://i.stack.imgur.com/iKHtN.png)

```
Sub TestAutoFilter()

    'this is your block of data with headers
    Dim rngDataAndHeader As Range
    Set rngDataAndHeader = Range("B2").CurrentRegion

    'this will knock off the header row if you want data only
    Dim rngData As Range
    Set rngData = Intersect(rngDataAndHeader, rngDataAndHeader.Offset(1))

    'autofilter
    rngDataAndHeader.AutoFilter Field:=2, Criteria1:=64

    'get the visible cells INCLUDING the header row
    Dim rngVisible As Range
    Set rngVisible = rngDataAndHeader.SpecialCells(xlCellTypeVisible)

    'check if there are more than 1 rows or if there are multiple areas (discontinuous range)
    If rngVisible.Rows.Count > 1 Or rngVisible.Areas.Count > 1 Then
        Debug.Print "found data"

        'data is available, this call cannot throw an error now
        Set rngVisible = rngData.SpecialCells(xlCellTypeVisible)

        'do your normal execution here
        '
        '
        '
    Else
        Debug.Print "only header, no data included"
    End If
End Sub

```

**Result with Criteria1:=64**

`Immediate window: found data`

[![enter image description here](https://i.stack.imgur.com/u6S0c.png)](https://i.stack.imgur.com/u6S0c.png)

**Result with Criteria1:=0**

`Immediate window: only header, no data included`

[![enter image description here](https://i.stack.imgur.com/xfiFe.png)](https://i.stack.imgur.com/xfiFe.png)

**Other notes:**

*   Code includes a separate variable called `rngData` if you want access to data without headers. This is just an INTERSECT-OFFSET to bump it one row down.
*   For the case where a result was found, code resets `rngVisible` to be the visible cells in the data only (skips header). Since this call cannot fail now, it is safe without error handling. This gives you a range that matches what you tried the first time but without the chance of getting an erorr. This is not required if you can process the original range `rngVisible` that includes the headers. If that is true, you can do away with `rngData` completely (unless you have some other need for it).
