```vb
Public Sub Chart_Axis_AutoX()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_Axis_AutoX
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Reverts the x axis of a chart back to Auto
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Dim targetChart As Chart
        Set targetChart = targetObject.Chart
        
        Dim xAxis As Axis
        Set xAxis = targetChart.Axes(xlCategory)
        xAxis.MaximumScaleIsAuto = True
        xAxis.MinimumScaleIsAuto = True
        xAxis.MajorUnitIsAuto = True
        xAxis.MinorUnitIsAuto = True

    Next targetObject

End Sub
```