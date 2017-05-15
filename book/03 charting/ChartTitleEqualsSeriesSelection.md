## ChartTitleEqualsSeriesSelection.md

```vb
Public Sub ChartTitleEqualsSeriesSelection()

    Dim targetObject As ChartObject

    For Each targetObject In Selection
        targetObject.Chart.ChartTitle.Text = targetObject.Chart.SeriesCollection(1).name
    Next targetObject
    
End Sub
```