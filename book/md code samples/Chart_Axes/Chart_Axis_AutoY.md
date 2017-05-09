```vb
Public Sub Chart_Axis_AutoY()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_Axis_AutoY
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Reverts the Y axis of a chart back to Auto
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Dim targetChart As Chart
        Set targetChart = targetObject.Chart
        
        Dim yAxis As Axis
        Set yAxis = targetChart.Axes(xlValue)
        yAxis.MaximumScaleIsAuto = True
        yAxis.MinimumScaleIsAuto = True
        yAxis.MajorUnitIsAuto = True
        yAxis.MinorUnitIsAuto = True

    Next targetObject

End Sub
```