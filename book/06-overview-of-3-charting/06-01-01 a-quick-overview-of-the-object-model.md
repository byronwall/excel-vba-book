### a quick overview of the object model

* ChartObjects -> ChartObject - this derives from `Shape` and exists when the Chart is on a Worksheet
    * Chart
        * SeriesCollection -> Series
        * Axes -> Axis
        * ChartArea
        * PlotArea
* ActiveChart -> Chart - this works whether you have a Worksheet or Chart on a sheet
* Selection -> Variant - this one can be useful but is often not of the type that you want.
