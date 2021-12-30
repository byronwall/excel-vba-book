## common objects/properties for a Chart

This section will focus on the common formatting changes that can be made to a Chart. The next section focuses on creating a Chart from scratch if you want to see that. These common changes will be grouped by the type that they affect, but this is not meant to be an exhaustive list. Instead, this is a list that will cover the objects nad functions that are actually used in regular code. There will be several other things that you will need to check the reference for (or record a macro), but this listing will get you started with the regular things.

To organize this section, we will focus on the different parts of a Chart in turn along with how to access the things you need. This section is meant to be a one stop shop for working on the common parts of a Chart. This will cover:

- `ChartObject`
  - Top, Left, Height, Width - control the location of a chart
- `Chart`
  - ChartType
  - Access the other objects and controls whether some things exist
    - HasLegend
    - HasTitle
- `Legend`
- `Series` -- accessed the the Chart.SeriesCollection
  - ChartType
- `Axis` -- accessed through Chart.Axes
  - Display the axis
  - Change the text
  - Change the min/max scale including automatic values
  - Change the number format of the axis
  - Change the format and display of the Gridlines
- `Point` -- accessed through a Series
  - Change display of individual points
  - Control the DataLabels (HasLabel and then DataLabel)
- `Trendline`

TODO: go through bUTL and find other commonly appearing things

### common changes to the ChartObject

The ChartObject is the main container for a Chart that is on a Worksheet. The common changes then are related to the position and size of the Chart on the Worksheet. The common properties to change here are:

- Top
- Left
- Height
- Width
- Placement (controls the move with cells option)

All of these are of type Double which means you can use decimal calculations to determine the size or position. In Excel, the 0,0 point is at the upper left hand corner (upper left of cell A1) and the Top and Left increase going to the right and down. If you are familiar with 0,0 being the center of the XY plane, then Excel will be a tad unfamiliar. Once you get used to it, you will realize that there is not really a better way to arrange the coordinate system since the spreadsheet can extend to the right and down nearly infinitely.

TODO: are there Bottom and Right properties too?

TODO: add a comment about Points vs. inches here and the function to convert them

The most common application of changing these properties is to either standardize the size of several charts or to arrange the charts in a grid (which standardizes the size and then position).

That code is included below:

TODO: clean up this code to only the required parts

```vb
Public Sub Chart_GridOfCharts( _
    Optional columnCount As Long = 3, _
    Optional chartWidth As Double = 400, _
    Optional chartHeight As Double = 300, _
    Optional offsetVertical As Double = 80, _
    Optional offsetHorizontal As Double = 40, _
    Optional shouldFillDownFirst As Boolean = False, _
    Optional shouldZoomOnGrid As Boolean = False)

    Dim targetObject As ChartObject

    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet

    Application.ScreenUpdating = False

    Dim countOfCharts As Long
    countOfCharts = 0

    For Each targetObject In targetSheet.ChartObjects
        Dim left As Double, top As Double

        If shouldFillDownFirst Then
            left = (countOfCharts \ columnCount) * chartWidth + offsetHorizontal
            top = (countOfCharts Mod columnCount) * chartHeight + offsetVertical
        Else
            left = (countOfCharts Mod columnCount) * chartWidth + offsetHorizontal
            top = (countOfCharts \ columnCount) * chartHeight + offsetVertical
        End If

        targetObject.top = top
        targetObject.left = left
        targetObject.Width = chartWidth
        targetObject.Height = chartHeight

        countOfCharts = countOfCharts + 1

    Next targetObject

    'loop through columns to find how far to zoom
    'Cells.Left property returns a variant in points
    If shouldZoomOnGrid Then
        Dim columnToZoomTo As Long
        columnToZoomTo = 1
        Do While targetSheet.Cells(1, columnToZoomTo).left < columnCount * chartWidth
            columnToZoomTo = columnToZoomTo + 1
        Loop

        targetSheet.Range("A:A", targetSheet.Cells(1, columnToZoomTo - 1).EntireColumn).Select
        ActiveWindow.Zoom = True
        targetSheet.Range("A1").Select
    End If

    Application.ScreenUpdating = True

End Sub
```

### common properties of the Chart

The Chart object is mostly a container for the other more useful properties of the Chart, but there are a couple of common properties that live at this top level. Those include:

- The HasXXX: HasTitle, HasLegend (TODO: any others?) - control the display of these things
- ChartType
- Delete
- Copy (TODO: this on ChartObject also?)

TODO: find more of these

In addition to those properties, the Chart object provides access to other useful things via the common accessors:

- SeriesCollection
- Axes
- Legend
- ChartTitle
- ChartArea
- PlotArea

TODO: is this list complete?

### common properties of the Series

One of the two most used Chart objects is the Series (other is the Axis). The Series ends up being powerful because it provides access to the data of the Chart along with the major formatting choices since the Series is the prominent feature of a Chart.

The common things to go after for a series are:

- Data related
  - Name
  - XValues
  - Values
  - Formula
- Formatting related
  - Format
    - Line
  - MarkerSize
  - MarkerStyle
  - MarkerForegroundColor, MarkerBackgroundColor

Also, from a Series you can access the following other objects:

- Points
- Trendlines

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

### common properties of the Legend

The Legend is a simple affair compared to the others. There really only two things to do with it: remove it or move it. Both of these are simple enough:

- HasLegend (on the Chart)
- Delete
- Position

TODO: add an example of these in action

### common properties of a Point

The Point represents the lowest level when it comes to how the data and formatting of a Chart is built. In general, you do not have to actively go editing Points. This is because you will typically edit the appearance of the Series and the Axes to get the Chart that you want. There are however times when you get down to the metal and edit the properties of the individual points. Before describing how to do this, it may help to give an example or two for why you want to get down to this level:

- Delete a data point without touching the Series
- Add a DataLabel to the point if the value is below some threshold (or if some other Range has a value)
- Hide a Point from one series because you want it to show up in another one

Of the tasks above, only one of them (the second) has to be accomplished via the Points. The others _could_ be done via a different method, but you might find yourself in a spot where iterating some Points will save a ton of headache elsewhere. A cautionary note is that typically you should not be editing the properties of a Point; there is nearly always a better way to do these things. Part of the problem is that the settings you change will be quickly overwritten by changes in Excel or VBA. If you know you just need something done however, Points can be a quick way to make it happen.

TODO: look into ErrorBars here?

WHen thinking about working through the Points of a Series, consider the common properties you can change:

- HasLabel / DataLabel
- Value
- Formatting? (TODO: what are these)
- Hidden

TODO: finish this list

Note that in addition to the common properties, you can also change anything that can be changed from the normal Excel settings/properties window.

### common properties of the TrendLine

The TrendLine is one of the lesser used properties, but it can be a real time saver when using VBA if you need to. The problem with the trendline normally is that you are required to work through a ton of menus to configure the properties. This is even more painful when you've got to do the same thing to multiple Series in a Chart or across multiple Charts. Similar to the other objects here, you can use VBA to quickly do the task that is otherwise a pain.

The most likely properties you'll use:

- Creating one off of a series
- Type
- Parameter

TODO: confirm these are correct

TODO: add an example showing how to add a Trendline for every Series

#### Chart_AddTrendlineToSeriesAndColor.md

```vb
Public Sub Chart_AddTrendlineToSeriesAndColor()

    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Dim chartIndex As Long
        chartIndex = 1

        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            'clear out old ones
            Dim j As Long
            For j = 1 To targetSeries.Trendlines.Count
                targetSeries.Trendlines(j).Delete
            Next j

            targetSeries.MarkerBackgroundColor = Chart_GetColor(chartIndex)

            Dim newTrendline As Trendline
            Set newTrendline = targetSeries.Trendlines.Add()
            newTrendline.Type = xlLinear
            newTrendline.Border.Color = targetSeries.MarkerBackgroundColor

            '2015 11 06 test to avoid error without name
            '2015 12 07 dealing with multi-cell Names
            'TODO: handle if the name is not a range also
            If Not butlSeries.name Is Nothing Then
                newTrendline.name = butlSeries.name.Cells(1, 1).Value
            End If

            newTrendline.DisplayEquation = True
            newTrendline.DisplayRSquared = True
            newTrendline.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Chart_GetColor(chartIndex)

            chartIndex = chartIndex + 1
        Next targetSeries

    Next targetObject
End Sub
```
