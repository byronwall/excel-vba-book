```vb
Public Sub ChartTitleEqualsSeriesSelection()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartTitleEqualsSeriesSelection
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets the chart title equal to the name of the first series
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Selection
        targetObject.Chart.ChartTitle.Text = targetObject.Chart.SeriesCollection(1).name
    Next targetObject
    
End Sub
```