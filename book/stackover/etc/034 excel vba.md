# SO item 034
I have a piece of code that is supposed to format a tab of information. I took it from a piece of code previously used and I am modifying it to fit my needs. I am getting a `syntax error` on the `.Selection Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(8, 10), _ Replace:=True, PageBreaks:=False, SummaryBelowData:=True` line.

VBA isn't exactly a strong point of mine so I would like to figure out a way to perform this properly. I know selection is frowned upon so if anyone has a way around it without me having to redo a bunch of code that would be awesome.

```
With ActiveSheet
.Range("A10").Select
.Range(Selection, Selection.End(xlToRight)).Select
.Range(Selection, Selection.End(xlDown)).Select
.Selection Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(8, 10), _
    Replace:=True, PageBreaks:=False, SummaryBelowData:=True
.Outline.ShowLevels RowLevels:=2

.Range("C8").Select
End With

```

----

Working code should be:

```
With ActiveSheet
.Range("A10").Select
.Range(Selection, Selection.End(xlToRight)).Select
.Range(Selection, Selection.End(xlDown)).Select
Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(8, 10), _
    Replace:=True, PageBreaks:=False, SummaryBelowData:=True
.Outline.ShowLevels RowLevels:=2

.Range("C8").Select
End With

```

`Subtotal` is a sub that needs to be called on the `Selection`. Therefore it needs a period between the two to make the call.

Also `Selection` is not a property of the `ActiveSheet` so the preceding period should be dropped inside the `With` block.
