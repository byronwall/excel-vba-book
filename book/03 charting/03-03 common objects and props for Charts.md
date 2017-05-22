## common objects/properties for a Chart

This section will focus on teh common formatting hcanges that can be made to a Chart.  Teh next section focuses on creatng a Chart from scratch if you want to see that. These common changes will be grouped by the type that they affect, but this is not meant to be an exhaustive list.  Instead, this is a list htat will ocver the objects nad functions that are actually used in regular code.  There will be several other things that you will need to check the reference for (or record a macro), but this listing will get you started with the regualrr things.

To organize this seciton, we will focus on the different parts of a Chart in turn along with how to access teh things you need.  This seciton is meant to be a one stop shop for working on teh common parts of a Chart.  This will cover:

* ChartObject
    * Top, Left, Height, Width - control the location of a chart
* Chart
    * ChartType
    * Access the other objects and controls whether some things exist
        * HasLegend
        * HasTitle
* Legend
* Series -- accessed the the Chart.SeriesCollection
    * ChartType
* Axis -- accessed through Chart.Axes
    * Display the axis
    * Change the text
    * Change the min/max scale including automatic values
    * Change the number format of the axis
    * Change the format and display of the Gridlines
* Point -- accessed through a Series
    * Change display of invidiual points
    * Control the DataLabels (HasLabel and then DataLabel)
* Trendline

TODO: go through bUTL and find other commonly appearing things
