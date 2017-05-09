```vb
Public Sub ChartMergeSeries()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartMergeSeries
    ' Author    : @byronwall
    ' Date      : 2015 12 30
    ' Purpose   : Merges all selected charts into a single chart
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    Dim targetChart As Chart
    Dim firstChart As Chart

    Dim isFirstChart As Boolean
    isFirstChart = True
    
    Application.ScreenUpdating = False
    
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
    
        Set targetChart = targetObject.Chart
        If isFirstChart Then
            Set firstChart = targetChart
            isFirstChart = False
        Else
            Dim targetSeries As series
            For Each targetSeries In targetChart.SeriesCollection

                Dim newChartSeries As series
                Dim butlSeries As New bUTLChartSeries

                butlSeries.UpdateFromChartSeries targetSeries
                Set newChartSeries = butlSeries.AddSeriesToChart(firstChart)

                newChartSeries.MarkerSize = targetSeries.MarkerSize
                newChartSeries.MarkerStyle = targetSeries.MarkerStyle

                targetSeries.Delete

            Next targetSeries

            targetObject.Delete

        End If
    Next targetObject
    
    Application.ScreenUpdating = True

End Sub
```