```vb
Public Sub ChartApplyToAll()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartApplyToAll
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Forces all charts to be a XYScatter type
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        targetObject.Chart.SeriesCollection(1).ChartType = xlXYScatter
    Next targetObject

End Sub
```