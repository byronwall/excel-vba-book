```vb
Public Sub Chart_AxisTitleIsSeriesTitle()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_AxisTitleIsSeriesTitle
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets the y axis title equal to the series name of the last series
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    Dim targetChart As Chart
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Set targetChart = targetObject.Chart

        Dim butlSeries As bUTLChartSeries
        Dim targetSeries As series

        For Each targetSeries In targetChart.SeriesCollection
            Set butlSeries = New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            targetChart.Axes(xlValue, targetSeries.AxisGroup).HasTitle = True
            targetChart.Axes(xlValue, targetSeries.AxisGroup).AxisTitle.Text = butlSeries.name

            '2015 11 11, adds the x-title assuming that the name is one cell above the data
            '2015 12 14, add a check to ensure that the XValue exists
            If Not butlSeries.XValues Is Nothing Then
                targetChart.Axes(xlCategory).HasTitle = True
                targetChart.Axes(xlCategory).AxisTitle.Text = butlSeries.XValues.Cells(1, 1).Offset(-1).Value
            End If

        Next targetSeries
    Next targetObject
End Sub
```