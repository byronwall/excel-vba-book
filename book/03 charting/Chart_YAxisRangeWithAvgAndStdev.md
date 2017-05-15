## Chart_YAxisRangeWithAvgAndStdev.md

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