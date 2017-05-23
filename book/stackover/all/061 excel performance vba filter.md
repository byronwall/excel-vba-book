# SO item 061
I would like to know if there is a faster way do this than the code I am using. I got the code using xlUp from the recorder.

```
 rCnt = Cells(Rows.Count, "B").End(xlUp).Row
 ActiveSheet.Range("$B$1:$J" & rCnt).AutoFilter Field:=5, _
      Criteria1:=Application.Transpose(arrCodes), Operator:=xlFilterValues
 Rows("2:" & rCnt).Delete Shift:=xlUp

```

And actually, if there was some way to flip the filter, I wouldn't need to delete at all as this is a temporary table that I copy from. However, all my research has failed to find a way to do

```
Criteria1:=Application.Transpose(<>arrCodes)

```

and arrCodes has too many elements to list in the filter. And the stuff that is not in arrCodes is way too numerous to make an array from. Thanks.

----

If you want to just use Excel UI and not formulas or VBA, you can do the following simple steps to get an "inverse" filter. This could then be ported to VBA if needed:

*   Apply the filter with the opposite conditions
*   Color those cells in one column (either font or background)
*   Clear the filter
*   Filter again but this time by cells in that column without color
*   Copy those results where you want them

This will not work well if the column already has some background colors. If that is the case, you can add a new column and color it. If this is in VBA, you could automate those steps. There are limits, but this is quick and simple if it applies.
