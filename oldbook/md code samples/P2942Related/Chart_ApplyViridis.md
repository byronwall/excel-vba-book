```vb
Public Sub Chart_ApplyViridis()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_ApplyTrendColors
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Applies the predetermined chart colors to each series
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            targetSeries.MarkerForegroundColorIndex = xlColorIndexNone
            targetSeries.MarkerStyle = xlMarkerStyleNone
            targetSeries.Format.Line.ForeColor.RGB = Chart_GetViridis(butlSeries.SeriesNumber, targetObject.Chart.SeriesCollection.Count)
            targetSeries.Format.Line.Weight = 1.5

        Next targetSeries
    Next targetObject
End Sub
```