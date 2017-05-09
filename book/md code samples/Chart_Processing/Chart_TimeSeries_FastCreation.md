```vb
Public Sub Chart_TimeSeries_FastCreation()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_TimeSeries_FastCreation
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : this will create a fast set of charts from a block of data
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim rangeOfDates As Range
    Dim dataRange As Range
    Dim rangeOfTitles As Range

    'dates are in B4 and down
    Set rangeOfDates = RangeEnd_Boundary(Range("B4"), xlDown)

    'data starts in C4, down and over
    Set dataRange = RangeEnd_Boundary(Range("C4"), xlDown, xlToRight)

    'titels are C2 and over
    Set rangeOfTitles = RangeEnd_Boundary(Range("C2"), xlToRight)

    Chart_TimeSeries rangeOfDates, dataRange, rangeOfTitles
    ChartDefaultFormat
    Chart_GridOfCharts

End Sub
```