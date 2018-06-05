# SO item 072
For a pivot table (pt1) on Sheet1, I use VBA to change the value of a filter field (filterfield) using the code below. Let's say values for field can be A, B or C

```
Sheets("Sheet1").PivotTables("pt1").PivotFields("filterfield").CurrentPage = "A"

```

Occasionally, and let's say randomonly for the purposes of this question, A, B or C will not be a valid selection for filterfield. When VBA attempts to change the field, it throws a run-time error. I want to avoid this.

How can I check if my values are valid for filterfield before I run the code above? I would like to avoid using On Error and VBA does not have try/catch functionality..

----

You can iterate through the `PivotItems` and check the `Name` against your test.

```
Sub CheckIfPivotFieldContainsItem()

    Dim pt As PivotTable
    Set pt = Sheet1.PivotTables(1)

    Dim test_val As Variant
    test_val = "59"

    Dim pivot_item As PivotItem
    For Each pivot_item In pt.PivotFields("C").PivotItems
        If pivot_item.Name = test_val Then
            Debug.Print "MATCHES"
        End If
    Next pi

End Sub

```

**Relevant data** shows that a match should exist and indeed it returns `MATCHES`.

![pivot data](https://i.stack.imgur.com/PR6PA.png)
