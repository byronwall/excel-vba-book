## ChartFlipXYValues.md

```vb
Public Sub ChartFlipXYValues()

    Dim targetObject As ChartObject
    Dim targetChart As Chart
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Set targetChart = targetObject.Chart

        Dim butlSeriesies As New Collection
        Dim butlSeries As bUTLChartSeries

        Dim targetSeries As series
        For Each targetSeries In targetChart.SeriesCollection
            Set butlSeries = New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            Dim dummyRange As Range

            Set dummyRange = butlSeries.Values
            Set butlSeries.Values = butlSeries.XValues
            Set butlSeries.XValues = dummyRange

            'need to change the series name also
            'assume that title is same offset
            'code blocked for now
            If False And Not butlSeries.name Is Nothing Then
                Dim rowsOffset As Long, columnsOffset As Long
                rowsOffset = butlSeries.name.Row - butlSeries.XValues.Cells(1, 1).Row
                columnsOffset = butlSeries.name.Column - butlSeries.XValues.Cells(1, 1).Column

                Set butlSeries.name = butlSeries.Values.Cells(1, 1).Offset(rowsOffset, columnsOffset)
            End If

            butlSeries.UpdateSeriesWithNewValues

        Next targetSeries

        ''need to flip axis labels if they exist
        ''three cases: X only, Y only, X and Y

        If targetChart.Axes(xlCategory).HasTitle And Not targetChart.Axes(xlValue).HasTitle Then

            targetChart.Axes(xlValue).HasTitle = True
            targetChart.Axes(xlValue).AxisTitle.Text = targetChart.Axes(xlCategory).AxisTitle.Text
            targetChart.Axes(xlCategory).HasTitle = False

        ElseIf Not targetChart.Axes(xlCategory).HasTitle And targetChart.Axes(xlValue).HasTitle Then
            targetChart.Axes(xlCategory).HasTitle = True
            targetChart.Axes(xlCategory).AxisTitle.Text = targetChart.Axes(xlValue).AxisTitle.Text
            targetChart.Axes(xlValue).HasTitle = False

        ElseIf targetChart.Axes(xlCategory).HasTitle And targetChart.Axes(xlValue).HasTitle Then
            Dim swapText As String

            swapText = targetChart.Axes(xlCategory).AxisTitle.Text

            targetChart.Axes(xlCategory).AxisTitle.Text = targetChart.Axes(xlValue).AxisTitle.Text
            targetChart.Axes(xlValue).AxisTitle.Text = swapText

        End If

        Set butlSeriesies = Nothing

    Next targetObject

End Sub
```
