## Chart_SortSeriesByName.md

```vb
Public Sub Chart_SortSeriesByName()
    'this will sort series by names
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        'uses a simple bubble sort but it works... shouldn't have 1000 series anyways
        Dim firstChartIndex As Long
        Dim secondChartIndex As Long
        For firstChartIndex = 1 To targetObject.Chart.SeriesCollection.Count
            For secondChartIndex = (firstChartIndex + 1) To targetObject.Chart.SeriesCollection.Count

                Dim butlSeries1 As New bUTLChartSeries
                Dim butlSeries2 As New bUTLChartSeries

                butlSeries1.UpdateFromChartSeries targetObject.Chart.SeriesCollection(firstChartIndex)
                butlSeries2.UpdateFromChartSeries targetObject.Chart.SeriesCollection(secondChartIndex)

                If butlSeries1.name.Value > butlSeries2.name.Value Then
                    Dim indexSeriesSwap As Long
                    indexSeriesSwap = butlSeries2.SeriesNumber
                    butlSeries2.SeriesNumber = butlSeries1.SeriesNumber
                    butlSeries1.SeriesNumber = indexSeriesSwap
                    butlSeries2.UpdateSeriesWithNewValues
                    butlSeries1.UpdateSeriesWithNewValues
                End If
                
            Next secondChartIndex
        Next firstChartIndex
    Next targetObject
End Sub
```