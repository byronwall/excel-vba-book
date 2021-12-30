## Chart_ApplyTrendColors.md

```vb
Public Sub Chart_ApplyTrendColors()

    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            targetSeries.MarkerForegroundColorIndex = xlColorIndexNone
            targetSeries.MarkerBackgroundColor = Chart_GetColor(butlSeries.SeriesNumber)

            targetSeries.Format.Line.ForeColor.RGB = targetSeries.MarkerBackgroundColor

        Next targetSeries
    Next targetObject
End Sub
```
