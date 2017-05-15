```vb
Public Sub Chart_ApplyFormattingToSelected()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_ApplyFormattingToSelected
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Applies a semi-random format to all charts
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
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