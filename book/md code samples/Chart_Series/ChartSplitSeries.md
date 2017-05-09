```vb
Public Sub ChartSplitSeries()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartSplitSeries
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Take all series from selected charts and puts them in their own charts
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    Dim targetChart As Chart

    Dim targetSeries As series
    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim newChartObject As ChartObject
            Set newChartObject = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)

            Dim newChartSeries As series
            Dim butlSeries As New bUTLChartSeries

            butlSeries.UpdateFromChartSeries targetSeries
            Set newChartSeries = butlSeries.AddSeriesToChart(newChartObject.Chart)

            newChartSeries.MarkerSize = targetSeries.MarkerSize
            newChartSeries.MarkerStyle = targetSeries.MarkerStyle

            targetSeries.Delete

        Next targetSeries


        targetObject.Delete

    Next targetObject
End Sub
```