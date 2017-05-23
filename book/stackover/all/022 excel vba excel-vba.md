# SO item 022
I have a sheet calls "Recap" and I want to know how much line that I have in this sheet.I tried with this code:

```
Function FindingLastRow(Mysheet As String) As Long

Dim sht As Worksheet
Dim LastRow As Long

Set sht = ThisWorkbook.Worksheets(Mysheet) 
LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
FindingLastRow = LastRow

End Function

```

......

in my macro i tried this:

....

```
Dim lastR As Long
lastR=FindingLastRow("Recap")
msgBox lastR

```

.....

----

The `UsedRange` on a Worksheet variable is very helpful here. You really don't need a UDF to get the row count.

```
LastRow = Worksheets("Recap").UsedRange.Rows.Count

```

**This method only works** if your data starts in row 1 and the sheet does not have formatting outside of the data. You could add in the starting row `+ UsedRange.Cells(1,1).Row` if you know the data starts somewhere other than row 1\. The second issue prevents the use of UsedRange.
