## Chart_AddTrendlineToSeriesAndColor.md

```vb
Public Sub Chart_AddTrendlineToSeriesAndColor()

    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Dim chartIndex As Long
        chartIndex = 1
        
        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            'clear out old ones
            Dim j As Long
            For j = 1 To targetSeries.Trendlines.Count
                targetSeries.Trendlines(j).Delete
            Next j

            targetSeries.MarkerBackgroundColor = Chart_GetColor(chartIndex)

            Dim newTrendline As Trendline
            Set newTrendline = targetSeries.Trendlines.Add()
            newTrendline.Type = xlLinear
            newTrendline.Border.Color = targetSeries.MarkerBackgroundColor
            
            '2015 11 06 test to avoid error without name
            '2015 12 07 dealing with multi-cell Names
            'TODO: handle if the name is not a range also
            If Not butlSeries.name Is Nothing Then
                newTrendline.name = butlSeries.name.Cells(1, 1).Value
            End If

            newTrendline.DisplayEquation = True
            newTrendline.DisplayRSquared = True
            newTrendline.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Chart_GetColor(chartIndex)

            chartIndex = chartIndex + 1
        Next targetSeries

    Next targetObject
End Sub
```