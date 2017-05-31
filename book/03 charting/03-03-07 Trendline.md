### common properties of the TrendLine

The TrendLine is one of the lesser used properties, but it can be a real time saver when using VBA if you need to.  The problem with the trendline normally is that you are required to work through a ton of menus to configure the properties.  This is even more painful when you've got to do the same thing to multiple Series in a Chart or across multiple Charts.  Similar to the other objects here, you can use VBA to quickly do the task that is otherwise a pain.

The most likely properties you'll use:

* Creating one off of a series
* Type
* Parameter

TODO: confirm these are correct
TODO: add an example showing how to add a Trendline for every Series

#### Chart_AddTrendlineToSeriesAndColor.md

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
