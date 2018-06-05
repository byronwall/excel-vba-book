## AddSeriesToChart.md

```vb
Public Function AddSeriesToChart(ByVal targetChart As Chart) As series

    Dim targetSeries As series
    Set targetSeries = targetChart.SeriesCollection.NewSeries
    
    targetSeries.Formula = Me.seriesFormula
    
    If Me.ChartType <> 0 Then targetSeries.ChartType = Me.ChartType
    
    Set AddSeriesToChart = targetSeries

End Function
```