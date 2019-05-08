## Chart_CreateDataLabels.md

```vb
Public Sub Chart_CreateDataLabels()

    Dim targetObject As ChartObject
    On Error GoTo Chart_CreateDataLabels_Error

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim dataPoint As Point
            Set dataPoint = targetSeries.Points(2)

            dataPoint.HasDataLabel = False
            dataPoint.DataLabel.Position = xlLabelPositionRight
            dataPoint.DataLabel.ShowSeriesName = True
            dataPoint.DataLabel.ShowValue = False
            dataPoint.DataLabel.ShowCategoryName = False
            dataPoint.DataLabel.ShowLegendKey = True

        Next targetSeries
    Next targetObject

    On Error GoTo 0
    Exit Sub

Chart_CreateDataLabels_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Chart_CreateDataLabels of Module Chart_Format"

End Sub
```
