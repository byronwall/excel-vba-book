### common changes to the ChartObject

The ChartObject is the main container for a Chart that is on a Worksheet.  The common changes then are related to the position and size of the Chart on the Worksheet.  The common properties to change here are:

* Top
* Left
* Height
* Width
* Placement (controls the move with cells option)

All of these are of type Double which means you can use decimal calculations to determine the size or position.  In Excel, the 0,0 point is at the upper left hand corner (upper left of cell A1) and the Top and Left increase going to the right and down.  If you are familiar with 0,0 being the center of the XY plane, then Excel will be a tad unfamiliar. Once you get used to it, you will realize that there is not really a better way to arrange the coordinate system since the spreadsheet can extend to the right and down nearly infinitely.

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
