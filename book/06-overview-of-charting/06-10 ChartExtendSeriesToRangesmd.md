## Chart_ExtendSeriesToRanges.md

```vb
Public Sub Chart_ExtendSeriesToRanges()

    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series

        'get each series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            'create the bUTL obj and manipulate series ranges
            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            If Not butlSeries.XValues Is Nothing Then
                targetSeries.XValues = RangeEnd(butlSeries.XValues.Cells(1), xlDown)
            End If
            targetSeries.Values = RangeEnd(butlSeries.Values.Cells(1), xlDown)

        Next targetSeries
    Next targetObject
End Sub
```
