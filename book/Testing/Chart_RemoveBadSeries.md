```vb
Public Sub Chart_RemoveBadSeries()
    
    'searches for bad charts on the current sheet
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(ActiveSheet)
        
        'iterate backwards since it's possible to delete a series from the chart
        Dim seriesIndex As Long
        For seriesIndex = targetObject.Chart.SeriesCollection.Count To 1 Step -1
            Dim seriesToCheck As series
            Set seriesToCheck = targetObject.Chart.SeriesCollection(seriesIndex)
            If Not SeriesHasFormula(seriesToCheck) Then
                Debug.Print "Series was removed " & seriesToCheck.name
                seriesToCheck.Delete
            End If
        Next
    Next targetObject
End Sub
```