## Chart_FitAxisToMaxAndMin.md

```vb
Public Sub Chart_FitAxisToMaxAndMin(ByVal axisType As XlAxisType)

    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        '2015 11 09 moved first inside loop so that it works for multiple charts
        Dim isFirst As Boolean
        isFirst = True

        Dim targetChart As Chart
        Set targetChart = targetObject.Chart

        Dim targetSeries As series
        For Each targetSeries In targetChart.SeriesCollection

            Dim minSeriesValue As Double
            Dim maxSeriesValue As Double

            If axisType = xlCategory Then

                minSeriesValue = Application.Min(targetSeries.XValues)
                maxSeriesValue = Application.Max(targetSeries.XValues)

            ElseIf axisType = xlValue Then

                minSeriesValue = Application.Min(targetSeries.Values)
                maxSeriesValue = Application.Max(targetSeries.Values)

            End If

            Dim targetAxis As Axis
            Set targetAxis = targetChart.Axes(axisType)

            Dim isNewMax As Boolean, isNewMin As Boolean
            isNewMax = maxSeriesValue > targetAxis.MaximumScale
            isNewMin = minSeriesValue < targetAxis.MinimumScale

            If isFirst Or isNewMin Then targetAxis.MinimumScale = minSeriesValue
            If isFirst Or isNewMax Then targetAxis.MaximumScale = maxSeriesValue

            isFirst = False
        Next targetSeries
    Next targetObject

End Sub
```