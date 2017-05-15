```vb
Public Function AddSeriesToChart(ByVal targetChart As Chart) As series
    '---------------------------------------------------------------------------------------
    ' Procedure : AddSeriesToChart
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Adds the represented series to a chart
    '---------------------------------------------------------------------------------------
    '
    Dim targetSeries As series
    Set targetSeries = targetChart.SeriesCollection.NewSeries
    
    targetSeries.Formula = Me.seriesFormula
    
    If Me.ChartType <> 0 Then targetSeries.ChartType = Me.ChartType
    
    Set AddSeriesToChart = targetSeries

End Function
```