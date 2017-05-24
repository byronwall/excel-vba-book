# SO item 025
I have a series of data which I'd like to plot into a few graphs. The data is ordered by date and extracted data tagged with a TRUE condition next to it based on prior conditions.

As there is a large data set, is it possible to generate a series of graphs automatically? I know how to create macros using the record macro function, just that I'm unsure how to filter the data to make the graphs.

In the sample data below, I would have 2 graphs with date ranges 1-3 Jan and another from 6-7 Jan.

E.g:

```
T/F   Date    Data
True  1-Jan   0.1
True  2-Jan   0.2
True  3-Jan   0.4
False 4-Jan   0.2
False 5-Jan   0.1
True  6-Jan   0.3
True  7-Jan   0.4

```

----

This is doable. Creating charts dynamically is one of those routines that you should put away in a library for future reference. I have done so and the code is below. The code will create a chart based on x/y ranges and a location. The location allows the charts to be arranged in a grid as they are created. You will have to wrangle your ranges so that you can give the sub below the inputs it needs. This should just be a matter of iterating through and tracking which charts to create.

The only key steps to this are using `ChartObjects.Add` to create a new chart (with positioning data) and then `SeriesCollection.NewSeries` to add a series to the chart.

You can call this code several times with `location` incrementing to create the charts you want and put them in the grid.

```
Sub CreateChartFromRange(xval As Range, yval As Range, location As Integer)

    Dim height As Double, width As Double
    height = 300
    width = 300

    Dim columns As Integer
    columns = 3

    'assume active sheet
    Dim cht_obj As ChartObject
    Set cht_obj = ActiveSheet.ChartObjects.Add( _
        (location Mod columns) * width, _
        (location \ columns) * height, _
        width, _
        height)

    Dim ser As Series
    Set ser = cht_obj.Chart.SeriesCollection.NewSeries

    ser.Values = yval
    ser.XValues = xval

    'assume XY scatter type
    ser.ChartType = xlXYScatter

End Sub

```
