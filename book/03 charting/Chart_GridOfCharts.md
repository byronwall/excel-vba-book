```vb
Public Sub Chart_GridOfCharts( _
    Optional columnCount As Long = 3, _
    Optional chartWidth As Double = 400, _
    Optional chartHeight As Double = 300, _
    Optional offsetVertical As Double = 80, _
    Optional offsetHorizontal As Double = 40, _
    Optional shouldFillDownFirst As Boolean = False, _
    Optional shouldZoomOnGrid As Boolean = False)
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GridOfCharts
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates a grid of charts.  Used by the form.
    '---------------------------------------------------------------------------------------
    '
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