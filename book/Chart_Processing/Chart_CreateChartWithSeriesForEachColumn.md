```vb
Public Sub Chart_CreateChartWithSeriesForEachColumn()
    'will create a chart that includes a series with no x value for each column

    Dim dataRange As Range
    Set dataRange = GetInputOrSelection("Select chart data")

    'create a chart
    Dim targetObject As ChartObject
    Set targetObject = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)
    
    targetObject.Chart.ChartType = xlXYScatter

    Dim targetColumn As Range
    For Each targetColumn In dataRange.Columns

        Dim chartDataRange As Range
        Set chartDataRange = RangeEnd(targetColumn.Cells(1, 1), xlDown)
        
        Dim butlSeries As New bUTLChartSeries
        Set butlSeries.Values = chartDataRange
        
        butlSeries.AddSeriesToChart targetObject.Chart
    Next targetColumn

End Sub
```