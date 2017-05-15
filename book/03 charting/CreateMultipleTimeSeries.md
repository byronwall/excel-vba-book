## CreateMultipleTimeSeries.md

```vb
Public Sub CreateMultipleTimeSeries()

    Dim rangeOfDates As Range
    Dim dataRange As Range
    Dim rangeOfTitles As Range

    On Error GoTo CreateMultipleTimeSeries_Error

    DeleteAllCharts

    Set rangeOfDates = Application.InputBox("Select date range", Type:=8)
    Set dataRange = Application.InputBox("Select data", Type:=8)
    Set rangeOfTitles = Application.InputBox("Select titles", Type:=8)

    Chart_TimeSeries rangeOfDates, dataRange, rangeOfTitles

    On Error GoTo 0
    Exit Sub

CreateMultipleTimeSeries_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & "), likely due to Range selection."

End Sub
```