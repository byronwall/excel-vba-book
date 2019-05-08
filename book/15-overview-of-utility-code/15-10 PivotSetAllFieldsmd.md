## PivotSetAllFields.md

```vb
Public Sub PivotSetAllFields()

    Dim targetTable As PivotTable
    Dim targetSheet As Worksheet

    Set targetSheet = ActiveSheet

    'this information is a bit unclear to me
    MsgBox "This defaults to the average for every Pivot table on the sheet.  Edit code for other result."
    On Error Resume Next
    For Each targetTable In targetSheet.PivotTables
        Dim targetField As PivotField
        For Each targetField In targetTable.DataFields
            targetField.Function = xlAverage
        Next targetField
    Next targetTable

End Sub
```
