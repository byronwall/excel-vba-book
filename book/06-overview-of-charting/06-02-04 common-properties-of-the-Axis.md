### common properties of the Axis

The Axis is the second most common object to work with (behind the Series). This is largely because the Axis controls or provides access to a lot of the formatting related aspects of the Chart. The Axis also controls the scale of the Axis and in that regard, is a critical part of making or editing a Chart.

The first part of the Axis is accessing the correct one. This is slightly tricky the first time because the Axes are stored in the Chart.Axes object. THe real trick is that this object is indexed by the xlAxisType (TODO: check that) which can be xlCategory (for the x-axis) or xlValue/xlValue2 (for the y-axis, left and right).

Once you have an Axis object, you can set to work changing the common properties:

- Scale related
  - MinimumScale/MaximumScale
  - MinimumScaleIsAuto/MaximumScaleIsAuto
- Formatting related (most of these are accessors to a different object)
  - GridLines (Major/minor and the HasXXX)
  - Ticks (TODO: that right?)
  - HasTitle and AxisTitle

#### Chart_Axis_AutoX.md

```vb
Public Sub Chart_Axis_AutoX()

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

#### Chart_Axis_AutoY.md

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

#### Chart_AxisTitleIsSeriesTitle.md

```vb
Public Sub Chart_AxisTitleIsSeriesTitle()

    Dim targetObject As ChartObject
    Dim targetChart As Chart
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Set targetChart = targetObject.Chart

        Dim butlSeries As bUTLChartSeries
        Dim targetSeries As series

        For Each targetSeries In targetChart.SeriesCollection
            Set butlSeries = New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            targetChart.Axes(xlValue, targetSeries.AxisGroup).HasTitle = True
            targetChart.Axes(xlValue, targetSeries.AxisGroup).AxisTitle.Text = butlSeries.name

            '2015 11 11, adds the x-title assuming that the name is one cell above the data
            '2015 12 14, add a check to ensure that the XValue exists
            If Not butlSeries.XValues Is Nothing Then
                targetChart.Axes(xlCategory).HasTitle = True
                targetChart.Axes(xlCategory).AxisTitle.Text = butlSeries.XValues.Cells(1, 1).Offset(-1).Value
            End If

        Next targetSeries
    Next targetObject
End Sub
```

#### Chart_FitAxisToMaxAndMin.md

```vb
Public Sub Chart_FitAxisToMaxAndMin(ByVal axisType As XlAxisType)

    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        '2015 11 09 moved first inside loop so that it works for multiple charts
        Dim isFirst As Boolean
        isFirst = True

        Dim targetChart As Chart
        Set targetChart = targetObject.Chart

        Dim targetSeries As series
        For Each targetSeries In targetChart.SeriesCollection

            Dim minSeriesValue As Double
            Dim maxSeriesValue As Double

            If axisType = xlCategory Then

                minSeriesValue = Application.Min(targetSeries.XValues)
                maxSeriesValue = Application.Max(targetSeries.XValues)

            ElseIf axisType = xlValue Then

                minSeriesValue = Application.Min(targetSeries.Values)
                maxSeriesValue = Application.Max(targetSeries.Values)

            End If

            Dim targetAxis As Axis
            Set targetAxis = targetChart.Axes(axisType)

            Dim isNewMax As Boolean, isNewMin As Boolean
            isNewMax = maxSeriesValue > targetAxis.MaximumScale
            isNewMin = minSeriesValue < targetAxis.MinimumScale

            If isFirst Or isNewMin Then targetAxis.MinimumScale = minSeriesValue
            If isFirst Or isNewMax Then targetAxis.MaximumScale = maxSeriesValue

            isFirst = False
        Next targetSeries
    Next targetObject

End Sub
```

#### Chart_YAxisRangeWithAvgAndStdev.md

```vb
Public Sub Chart_YAxisRangeWithAvgAndStdev()

    Dim numberOfStdDevs As Double

    numberOfStdDevs = CDbl(InputBox("How many standard deviations to include?"))

    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        Set targetSeries = targetObject.Chart.SeriesCollection(1)

        Dim avgSeriesValue As Double
        Dim stdSeriesValue As Double

        avgSeriesValue = WorksheetFunction.Average(targetSeries.Values)
        stdSeriesValue = WorksheetFunction.StDev(targetSeries.Values)

        targetObject.Chart.Axes(xlValue).MinimumScale = avgSeriesValue - stdSeriesValue * numberOfStdDevs
        targetObject.Chart.Axes(xlValue).MaximumScale = avgSeriesValue + stdSeriesValue * numberOfStdDevs

    Next

End Sub
```
