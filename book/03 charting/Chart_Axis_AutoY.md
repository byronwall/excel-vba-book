## Chart_Axis_AutoY.md

```vb
Public Sub Chart_Axis_AutoY()

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