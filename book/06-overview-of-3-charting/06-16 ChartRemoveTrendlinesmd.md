## Chart_RemoveTrendlines.md

```vb
Public Sub Chart_RemoveTrendlines()

    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim newTrendline As Trendline
            For Each newTrendline In targetSeries.Trendlines
                newTrendline.Delete
            Next newTrendline
        Next targetSeries
    Next targetObject
End Sub
```