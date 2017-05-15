```vb
Public Sub Chart_TrendlinesToAverage()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_TrendlinesToAverage
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates a trendline using a moving average instead of linear
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series

        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim newTrendline As Trendline

            For Each newTrendline In targetSeries.Trendlines
                newTrendline.Type = xlMovingAvg
                newTrendline.Period = 15
                newTrendline.Format.Line.Weight = 2
            Next
        Next
    Next

End Sub
```