```vb
Public Sub Chart_ExtendSeriesToRanges()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_ExtendSeriesToRanges
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Extends the underlying data for a series to go to the end of its current Range
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series

        'get each series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            'create the bUTL obj and manipulate series ranges
            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            If Not butlSeries.XValues Is Nothing Then
                targetSeries.XValues = RangeEnd(butlSeries.XValues.Cells(1), xlDown)
            End If
            targetSeries.Values = RangeEnd(butlSeries.Values.Cells(1), xlDown)

        Next targetSeries
    Next targetObject
End Sub
```