```vb
Public Sub Chart_RemoveTrendlines()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_RemoveTrendlines
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Remove all trendlines from a chart
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim newTrendline As Trendline
            For Each newTrendline In targetSeries.Trendlines
                newTrendline.Delete
            Next newTrendline
        Next targetSeries
    Next targetObject
End Sub
```