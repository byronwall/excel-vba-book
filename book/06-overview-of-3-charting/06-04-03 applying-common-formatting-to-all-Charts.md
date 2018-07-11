### applying common formatting to all Charts

#### ChartDefaultFormat.md

```vb
Public Sub ChartDefaultFormat()

    Const MARKER_SIZE As Long = 3
    Dim majorGridlineColor As Long
    majorGridlineColor = RGB(242, 242, 242)
    Const TITLE_FONT_SIZE As Long = 12
    Const SERIES_LINE_WEIGHT As Single = 1.5

    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Dim targetChart As Chart

        Set targetChart = targetObject.Chart

        Dim targetSeries As series
        For Each targetSeries In targetChart.SeriesCollection

            targetSeries.MarkerSize = MARKER_SIZE
            targetSeries.MarkerStyle = xlMarkerStyleCircle

            If targetSeries.ChartType = xlXYScatterLines Then targetSeries.Format.Line.Weight = SERIES_LINE_WEIGHT

            targetSeries.MarkerForegroundColorIndex = xlColorIndexNone
            targetSeries.MarkerBackgroundColorIndex = xlColorIndexAutomatic

        Next targetSeries


        targetChart.HasLegend = True
        targetChart.Legend.Position = xlLegendPositionBottom

        Dim targetAxis As Axis
        Set targetAxis = targetChart.Axes(xlValue)

        targetAxis.MajorGridlines.Border.Color = majorGridlineColor
        targetAxis.Crosses = xlAxisCrossesMinimum

        Set targetAxis = targetChart.Axes(xlCategory)

        targetAxis.HasMajorGridlines = True

        targetAxis.MajorGridlines.Border.Color = majorGridlineColor

        If targetChart.HasTitle Then
            targetChart.ChartTitle.Characters.Font.Size = TITLE_FONT_SIZE
            targetChart.ChartTitle.Characters.Font.Bold = True
        End If

        Set targetAxis = targetChart.Axes(xlCategory)

    Next targetObject

End Sub
```

