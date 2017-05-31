## specific charting examples

This section will focus on some specific applications of applying VBA to charts.  The code here can be quickly reused for your own application.  These examples include:

* Creating a grid of XY scatter plots (a scatter matrix) based on a block of data
* Creating a panel of time series, one chart per each value with a common x-axis

TODO: identify the examples to include here

### creating an XY scatter matrix

#### ChartCreateXYGrid.md

```vb
Public Sub ChartCreateXYGrid()

    On Error GoTo ChartCreateXYGrid_Error

    DeleteAllCharts
    'VBA doesn't allow a constant to be defined using a function (rgb) so we use a local variable rather than
    'muddying it up with the calculated value of the rgb function
    Dim majorGridlineColor As Long
    majorGridlineColor = RGB(200, 200, 200)
    Dim minorGridlineColor As Long
    minorGridlineColor = RGB(220, 220, 220)

    Const CHART_HEIGHT As Long = 300
    Const CHART_WIDTH As Long = 400
    Const MARKER_SIZE As Long = 3
    'dataRange will contain the block of data with titles included
    Dim dataRange As Range
    Set dataRange = Application.InputBox("Select data with titles", Type:=8)

    Application.ScreenUpdating = False

    Dim rowIndex As Long, columnIndex As Long
    rowIndex = 0

    Dim xAxisDataRange As Range, yAxisDataRange As Range
    For Each yAxisDataRange In dataRange.Columns
        columnIndex = 0

        For Each xAxisDataRange In dataRange.Columns
            If rowIndex <> columnIndex Then
                Dim targetChart As Chart
                Set targetChart = ActiveSheet.ChartObjects.Add(columnIndex * CHART_WIDTH, _
                                                               rowIndex * CHART_HEIGHT + 100, _
                                                               CHART_WIDTH, CHART_HEIGHT).Chart

                Dim targetSeries As series
                Dim butlSeries As New bUTLChartSeries

                'offset allows for the title to be excluded
                Set butlSeries.XValues = Intersect(xAxisDataRange, xAxisDataRange.Offset(1))
                Set butlSeries.Values = Intersect(yAxisDataRange, yAxisDataRange.Offset(1))
                Set butlSeries.name = yAxisDataRange.Cells(1)
                butlSeries.ChartType = xlXYScatter

                Set targetSeries = butlSeries.AddSeriesToChart(targetChart)

                targetSeries.MarkerSize = MARKER_SIZE
                targetSeries.MarkerStyle = xlMarkerStyleCircle

                Dim targetAxis As Axis
                Set targetAxis = targetChart.Axes(xlCategory)
                targetAxis.HasTitle = True
                targetAxis.AxisTitle.Text = xAxisDataRange.Cells(1)
                targetAxis.MajorGridlines.Border.Color = majorGridlineColor
                targetAxis.MinorGridlines.Border.Color = minorGridlineColor

                Set targetAxis = targetChart.Axes(xlValue)
                targetAxis.HasTitle = True
                targetAxis.AxisTitle.Text = yAxisDataRange.Cells(1)
                targetAxis.MajorGridlines.Border.Color = majorGridlineColor
                targetAxis.MinorGridlines.Border.Color = minorGridlineColor

                targetChart.HasTitle = True
                targetChart.ChartTitle.Text = yAxisDataRange.Cells(1) & " vs. " & xAxisDataRange.Cells(1)
                'targetChart.ChartTitle.Characters.Font.Size = 8
                targetChart.Legend.Delete
            End If

            columnIndex = columnIndex + 1
        Next xAxisDataRange

        rowIndex = rowIndex + 1
    Next yAxisDataRange

    Application.ScreenUpdating = True

    dataRange.Cells(1, 1).Activate

    On Error GoTo 0
    Exit Sub

ChartCreateXYGrid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
           ") in procedure ChartCreateXYGrid of Module Chart_Format"
    MsgBox "This is most likely due to Range issues"

End Sub
```

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
