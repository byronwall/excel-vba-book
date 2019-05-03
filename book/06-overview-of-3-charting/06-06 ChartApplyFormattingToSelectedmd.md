## Chart_ApplyFormattingToSelected.md

```vb
Public Sub Chart_ApplyFormattingToSelected()

    Dim targetObject As ChartObject
    Const MARKER_SIZE As Long = 5

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series

        For Each targetSeries In targetObject.Chart.SeriesCollection
            targetSeries.MarkerSize = MARKER_SIZE
        Next targetSeries
    Next targetObject

End Sub
```
