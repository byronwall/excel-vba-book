### obtaining a reference to a Chart

When working with CHarts, the first task is typically to get a reference to an existing chart -- unless you are creating a new chart.  To obtain a reference to a chart, there are a handful of ways of doing it depending on what your spreadsheet contains and how it's structured.

THe main ways to do it are:

* Use the `ActiveChart` object
* Use the `Selection` object -- this is highly depending on what is selected
* Use the `ChartObjects` object
    * If you know which chart you want, you can supply an index; this works great if there is only a single chart - `ChartObjects(1)`
    * If you want to do something to all charts, you can iterate this object
    * If you have named the chart (more on that later) you can supply the name as the index - `ChartObjects("SomeChart")`
* The `Workbook.Sheets` object if your charts are contained in their own sheets
    * Same as above, you can access via a numeric index, name, or iterate through all of them

#### `ActiveChart`

`ActiveChart` is similar to the other `Active` objects in that it does about what you expect.  The one difference is that the Chart actually has to be selected or have focus in order to be considered "active".  This is similar but also different to something like `ActiveWorkbook` where having the workbook open makes it active.

Note that ActiveChart will work for a Chart that is contained on a Worksheet or also for one that is its own Sheet.  If the latter case, then ActiveSheet and ActiveChart will refer to the same object.  Side note: this technicality is why you will not get proper Intellisense when using ActiveSheet -- that Sheet could technically be a Chart.

The nice thing about ActiveChart is that it gives you the Chart object which then gives you immediate access to the Chart related details you are like to want to change.  The downside is that unless you have a single Chart that is already selected, ActiveChart has limited application when using VBA.  Again, the goal is to avoid selecting objects in order to access them via VBA so ActiveChart has this limitation.

#### `Selection`

The Selection object is probably the greatest catch all for an object.  It literally holds anything, and this means that using the object requires knowing what is selected, or checking vigorously before using the object.  Technically, you also let your code error out if the wrong object is selected, and this works well at times.  This works well because oyu are unlikely to be using Selection in a complicated workflow because, again, you should not be selecting objects to access them.  This means that Selection is really limited to one-off and helper code where you can more tightly dictate that this code only works if you select a Chart.  You should still add some error handling, but sometimes that step is skipped.

Since the Selection can hold anything, it's important to know what could be Selected.  Related to charts, the following can all live in the Selection:

* ChartObjects
* Chart
* ChartArea
* PlotArea
* Legend
* ChartTitle
* Series

If you are writing VBA to work on Charts, you can technically require the user to select the correct part of the chart and always use `Selection`.  You will quickly grow tired of having to remember which part of the Chart to select in order ot make the code work.  To avoid this scenario, it is helpful to remember the object model and know how to work your way around a Chart.

My approach has always been to convert the Selection to a Collection of ChartObjects. I can then always iterate that resulting Collection to process the Charts.  If only a single Chart was selected, the code works all the same.  The downside to this approach is that a Chart as a Sheet cannot live inside a ChartObject.  This is a large part of why I always put Charts on a Worksheet.

Below is the helper function I use in order to convert a possibly Chart containing selection into a Collection of ChartObjects.  It works for all objects except for the Axis related ones.

TODO: consider improving this code if it is included as a de facto reference

```vb
Public Function Chart_GetObjectsFromObject(ByVal inputObject As Object) As Variant

    Dim chartObjectCollection As New Collection

    'NOTE that this function does not work well with Axis objects.  Excel does not return the correct Parent for them.

    Dim targetObject As Variant
    Dim inputObjectType As String
    inputObjectType = TypeName(inputObject)

    Select Case inputObjectType

        Case "DrawingObjects"
            'this means that multiple charts are selected
            For Each targetObject In inputObject
                If TypeName(targetObject) = "ChartObject" Then
                    'add it to the set
                    chartObjectCollection.Add targetObject
                End If
            Next targetObject

        Case "Worksheet"
            For Each targetObject In inputObject.ChartObjects
                chartObjectCollection.Add targetObject
            Next targetObject

        Case "Chart"
            chartObjectCollection.Add inputObject.Parent

        Case "ChartArea", "PlotArea", "Legend", "ChartTitle"
            'parent is the chart, parent of that is the chart targetObject
            chartObjectCollection.Add inputObject.Parent.Parent

        Case "Series"
            'need to go up three levels
            chartObjectCollection.Add inputObject.Parent.Parent.Parent

        Case "Axis", "Gridlines", "AxisTitle"
            'these are the oddly unsupported objects
            MsgBox "Axis/gridline selection not supported.  This is an Excel bug.  Select another element on the chart(s)."

        Case Else
            MsgBox "Select a part of the chart(s), except an axis."

    End Select

    Set Chart_GetObjectsFromObject = chartObjectCollection
End Function
```

#### ChartObjects

If you are working on a Worksheet, then that Worksheet will have the ChartObjects object.  This object is great because it contains all of the Charts in their own collection (separate from any other Shapes or buttons).  This ChartObjects collection contains object of type ChartObject.  The ChartObject derives from Shape which means it contains all of the properties related to on-sheet position and size.

A typical workflow is included below since it is a pattern that shows up all the time in VBA code related to charts.  At a high level the steps are:

* Use ActiveSheet or a Worksheet reference to access the ChartObjects
* Iterate through each ChartObject, storing a reference to the underlying Chart
* You then setup sections to work through the parts of the Chart you want
    * Iterate through the SeriesCollection
    * Iterate through the Axes
    * Touch the other top level properties including ChartTile, Legend, etc.

This workflow is quite powerful because it can quickly be wrapped with a loop to go through all Worksheets and even possible all Workbooks.  It's also powerful because you can be quite comfortable learning this pattern and then adding in the parts that you actually want ot change.  The only downside is that it can be quite tedious to type out all the loops every time, but there's not a good way around that other than to use the clipboard.

Another approach to using ChartObjects is to not iterate through all of them but instead to select a single ChartObject and work with it.  There are two ways to do this:

* Use an integer index for the Chart -- this is quite easy to do if there are only a few charts
* Name the chart and use that name

When using either of these approaches, it is quite helpful to show the `Selection Pane` window in Excel.  This pane will pop out and tell you the order and the names of all the objects on the sheet (this includes comments, shapes, and Charts).  From this pane, you can rearrange the charts into the order you want or rename them.

Although `For Each` loops are generally preferred when working with Charts, sometimes you simply know that you want to change one chart and an index just lets you do that.  If you are in the habit of using loops however, you can easily do that with the helper code included above which stick a single chart into a Collection.

#### Workbook.Sheets to get Chart references

The final approach to obtaining a Chart reference is to use the Sheets object.  Aside from ActiveChart, this is the only way to deal with Charts that are their own Sheet.  Again, you cna either use an index or a Name.  Here, the Name is easily changed on the Sheet tab so it's much more common to use a Name when doing this.  The other approach is to iterate through all the Sheets and pick off the ones that are Charts.

There are two key points when working with Charts as Sheets:

* You must use the Workbook.Sheets object to access them and not Workbook.Worksheets.  The latter object contains only those Worksheets that are not Charts.  The former contains both Charts and Worksheets.
* It's possible that your Sheet is not actually a Chart.  You should check the type of the object is you are going to iterate through all Worksheets.  Also be aware that some sheets can be hidden which might lead to unexpected results.

TODO: is there a Charts object on Workbook?
