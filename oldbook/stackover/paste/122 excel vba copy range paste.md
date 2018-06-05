# SO item 122
What I want to do: Filter on an array, then copy the filtered data onto a different workbook. Then copy that data that was just pasted onto another worksheet in the same workbook, but this time below existing data.

My thinking: The code I used below was used to help me copy and paste from one workbook to another workbook and worked perfectly,

```
wb.Sheets("2014 Current Week").Range("C2:CC10000").Copy nwb.Sheets("2014 YTD").Range("C" & Rows.Count).End(xlUp).Offset(1, 0) 

```

but it seems it does not work the same way if I want to copy in the same workbook. Any help would be appreciated. THANKS!

```
Dim wb as workbook
Dim strs As String
Dim str As String
Dim nwb as workbook

Set wb = ThisWorkbook

strs = wb.Sheets("Macros").Range("H5") 'the 2014 address can be found in full in cell H5 in the Macros tab

set nwb = Workbooks.Open(strs) 'address of new workbook and opens it

With ActiveSheet

.AutofilterMode = False
'Filter this and that here'

End With

 nwb.Sheets("ALL DATA").Range("A1:CA100000").Copy wb.Sheets("2014 Current Week").Range("C" & Rows.Count).End(xlUp) 
'this one works, and copies exactly as I want into the 2014 Current Week tab

 wb.Sheets("2014 Current Week").Range("C2:CC10000").Copy wb.Sheets("2014 YTD").Range("C" & Rows.Count).End(xlUp).Offset(1, 0) 
'this one doesn't work, and does not copy or paste at all from that 2014 Current Week into the 2014 YTD tab of the same workbook

```

----

Your code looks good and should not have any issues. Since it seems to not be working, you need to look at the actual `Worksheets` and data involved.

You are using `End` to find the last row of data. Given this, it is worth testing this yourself. Go to the last row of the `Worksheet` in column `C` and hit <kbd>CTRL + UP</kbd>. This will show you where the data is going to be pasted.

Based on your description this last row is the issue. Since there was some stray data that was affecting `End`.
