# SO item 001
Chart = xy Scatter

I have a chart on which I have set a Horizontal Axis Crosses-Axis value = 200\. And that chart looks as follows:

![enter image description here](https://i.stack.imgur.com/ISljx.png)

Now I wish to reverse the order of the y-axis labels. But when I check "Values in reverse order" the x-axis labels jump above the x-axis line as follows:

![enter image description here](https://i.stack.imgur.com/UbxpI.png)

I'd like the x-axis labels to be bellow the x-axis line. I have been trying various VBA code and just various excel options. Such as:

.TickLabels.Offset (this offsets the x-axis labels further away from the x-axis, will not accept negative numbers)

I wanted to add a large MarginTop value to the x-axis labels, but I was unable to figure out the VBA code and the option is greyed out in Excel.

Any possible thoughts or solutions would be very much appreciated.

----

If you can live with the x-axis being at the bottom of the chart, you can get the labels back "under" the axis by changing the TickLabelPosition to be equal to xlTickLabelPositionHigh. It looks better if you also cross the y-axis at the "Maximum axis value". This puts the axis formatting at the bottom (really the maximum value) with the labels.

You can get all these settings from the normal menus. If you need VBA to do this, here is a starting point:

```
Sub reverseAxisAndLabelAtBottom()

    Dim cht As Chart
    Dim x_axis As Axis
    Dim y_axis As Axis

    'using the ActiveChart... assumes it is selected

    Set cht = ActiveChart
    Set x_axis = cht.Axes(xlCategory)
    Set y_axis = cht.Axes(xlValue)

    y_axis.ReversePlotOrder = True
    x_axis.TickLabelPosition = xlTickLabelPositionHigh
    y_axis.Crosses = xlAxisCrossesMaximum

End Sub

```

I also tried to do this by adding a dummy series and putting reversed y-axis on the secondary axis. This gets the labels right, but then the y-axis is forced over to the far right. I would not call that an "ideal" way of doing it.

Before/after pictures with some random data. ![before](https://i.stack.imgur.com/mcL96.png) ![after](https://i.stack.imgur.com/FUHiy.png)
