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
