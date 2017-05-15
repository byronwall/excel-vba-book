## ChartApplyToAll.md

```vb
Public Sub ChartApplyToAll()

    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        targetObject.Chart.SeriesCollection(1).ChartType = xlXYScatter
    Next targetObject

End Sub
```