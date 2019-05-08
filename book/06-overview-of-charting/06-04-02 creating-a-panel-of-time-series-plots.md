### creating a panel of time series plots

#### Chart_TimeSeries.md

```vb
Public Sub Chart_TimeSeries(ByVal rangeOfDates As Range, ByVal dataRange As Range, ByVal rangeOfTitles As Range)

    Application.ScreenUpdating = False
    Const MARKER_SIZE As Long = 3
    Dim majorGridlineColor As Long
    majorGridlineColor = RGB(200, 200, 200)
    Dim chartIndex As Long
    chartIndex = 1

    Dim titleRange As Range
    Dim targetColumn As Range

    For Each titleRange In rangeOfTitles

        Dim targetObject As ChartObject
        Set targetObject = ActiveSheet.ChartObjects.Add(chartIndex * 300, 0, 300, 300)

        Dim targetChart As Chart
        Set targetChart = targetObject.Chart
        targetChart.ChartType = xlXYScatterLines
        targetChart.HasTitle = True
        targetChart.Legend.Delete

        Dim targetAxis As Axis
        Set targetAxis = targetChart.Axes(xlValue)
        targetAxis.MajorGridlines.Border.Color = majorGridlineColor

        Dim targetSeries As series
        Dim butlSeries As New bUTLChartSeries

        Set butlSeries.XValues = rangeOfDates
        Set butlSeries.Values = dataRange.Columns(chartIndex)
        Set butlSeries.name = titleRange

        Set targetSeries = butlSeries.AddSeriesToChart(targetChart)

        targetSeries.MarkerSize = MARKER_SIZE
        targetSeries.MarkerStyle = xlMarkerStyleCircle

        chartIndex = chartIndex + 1

    Next titleRange

    Application.ScreenUpdating = True
End Sub
```
