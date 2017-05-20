## introduction to charting

Charting is the second most important aspect of automatic Excel behind manipulating `Ranges`.  There is a bias when saying that because a lot of what I do after engineering calculations is chart the results.  In particular, Excel can be used to great effect to chart time series of data.  THe other reason charts are so amenable to VBA is that very oftne you are applying the same actiosn to the charts.  In that sense, the VBA related ot charts is doing a lot of changing settings and formats so that the charts look the way you want. This ahs teh immediate effect of making your charts look less liek "they came from Excel" which si a common knock in some circles.

When workign with `Charts`, there is a `Range` of difficulties depending on what you are trying to do.  In some cases, workign with an exsiting `chart` is much easier than creating a new one.  In other instances, it can be much simpler to create a new chart rather, starting from adefault, rather than cahnge all the settings back.  One other major difference between `Charts` nd `Ranges` is that workign with charts is much more about knoiwng the object model than knowing how to program.  The vast majority of your code related to charts is ismple iterating through objects to find the one property that you want to change.  THis makes it easier to write chart VBA once you have the bascis of `For Each` loops down.  It also means that you need to spend some time getting comfortable woth the object model.

There is one oddity related to Charts that si wotht mentioning now.  Charts can either be embeeded as an objecy on a `Workseeht`, or they can be their own `Sheets`.  I personally never use the latter case, but it is common enough that it needs to be on your mind when working with Charting code.

(I don't use the Chart as a Sheet model because I find that it is not necessary in terms of displaying data.  In particular, you are at the mercy of your window size and cannot easily change the dimensions.  Also, it complicates the VBA side of things to work in both formats all the time, so I just decided to always put my CHarts on Sheets.  Your mileage may vary so I'll touch on both approaches in teh code samples.)

### a quick overview of the object model

* ChartObjects -> ChartObject - this derives from `Shape` and exists when the Chart is on a Worksheet
    * Chart
        * SeriesCollection -> Series
        * Axes -> Axis
        * ChartArea
        * PlotArea
* ActiveChart -> Chart - this works whether you have a Worksheet or Chart on a sheet
* Selection -> Variant - this one can be useful but is often not of the tpye that you want.
