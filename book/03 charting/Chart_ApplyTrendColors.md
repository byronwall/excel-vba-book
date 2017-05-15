```vb
Public Sub Chart_ApplyTrendColors()
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
            targetSeries.MarkerBackgroundColor = Chart_GetColor(butlSeries.SeriesNumber)

            targetSeries.Format.Line.ForeColor.RGB = targetSeries.MarkerBackgroundColor

        Next targetSeries
    Next targetObject
End Sub
```