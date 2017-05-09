```vb
Public Sub RemoveZeroValueDataLabel()
    '---------------------------------------------------------------------------------------
    ' Procedure : RemoveZeroValueDataLabel
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Code deletes data labels that have 0 value
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    'uses the ActiveChart, be sure a chart is selected
    Dim targetChart As Chart
    Set targetChart = ActiveChart

    Dim targetSeries As series
    For Each targetSeries In targetChart.SeriesCollection

        Dim seriesValues As Variant
        seriesValues = targetSeries.Values

        'include this line if you want to reestablish labels before deleting
        targetSeries.ApplyDataLabels xlDataLabelsShowLabel, , , , True, False, False, False, False

        'loop through values and delete 0-value labels
        Dim pointIndex As Long
        For pointIndex = LBound(seriesValues) To UBound(seriesValues)
            If seriesValues(pointIndex) = 0 Then
                With targetSeries.Points(pointIndex)
                    If .HasDataLabel Then .DataLabel.Delete
                End With
            End If
        Next pointIndex
    Next targetSeries
End Sub
```