```vb
Sub RemoveAllLegends()

    Dim chtObj As ChartObject
    
    For Each chtObj In Chart_GetObjectsFromObject(Selection)
        chtObj.Chart.HasLegend = False
        chtObj.Chart.HasTitle = True
        
        chtObj.Chart.SeriesCollection(1).MarkerSize = 4
    Next

End Sub
```