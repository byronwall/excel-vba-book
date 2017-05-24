# SO item 123
I'm currently having an unexpected issue with my code. It's suppose to update two data series of a chart object dynamically. Recently I've added another series to the collection(a total of 3 series now). The data series update properly but the problem is that the formatting for the 3 series now swap around between each other every time an update occurs. Here is my code below where the updating occurs:

```
Dim WrkSheet As String
WrkSheet = Application.Worksheets(X).Name

If Background > 0 Then
    'Update the background data series
    Application.Worksheets(X).ChartObjects("Main Chart").Chart.SeriesCollection(2).Formula = _
    "=SERIES(""Background"",('" & WrkSheet & "'!$J$32:$J$" & (32 + Background - 1) & ",'" & WrkSheet & "'!$J$" & BackgroundStart2 & ":$J$" & (BackgroundStart2 + Background - 1) & ")," _
    & "('" & WrkSheet & "'!$K$32:$K$" & (32 + Background - 1) & ",'" & WrkSheet & "'!$K$" & BackgroundStart2 & ":$K$" & (BackgroundStart2 + Background - 1) & "),2)"
Else
    'Make the application not graph background in this scenario
End If

'Update the Peak data series
Application.Worksheets(X).ChartObjects("Main Chart").Chart.SeriesCollection(1).Formula = _
"=SERIES(""Peak"",'" & WrkSheet & "'!$J$" & (PeakStart1) & ":$J$" & PeakEnd1 & ",'" & WrkSheet & "'!$K$" & PeakStart1 & ":$K$" & PeakEnd1 & ",1)"

'Update the peak background data series
Application.Worksheets(X).ChartObjects("Main Chart").Chart.SeriesCollection(3).Formula = _
"=SERIES(""Step Background"",'" & WrkSheet & "'!$J$" & (PeakStart1) & ":$J$" & PeakEnd1 & ",'" & WrkSheet & "'!$O$" & PeakStart1 & ":$O$" & PeakEnd1 & ",1)"

```

Once this code completes, each of the 3 series collection objects update correctly but the associated formatting for each changes. I believe that the series collection may be deleted and recreated removing the formatting, but I'm unsure why this would be the case. Any help would be great.

----

The final parameter of the `SERIES` call is the index order. You have two entries with a 1 and the first one has a 2\. They are probably displacing each other as you go. You should number those in order (same order as their spot in the `SeriesCollection`).

**Code** shows changing the last formula which appears to be the errant one.

```
"=SERIES(""Step Background"",'" & WrkSheet & "'!$J$" & (PeakStart1) & ":$J$" & PeakEnd1 & ",'" & WrkSheet & "'!$O$" & PeakStart1 & ":$O$" & PeakEnd1 & ",3)"

```

Note that I changed the last row from a `1` to a `3` to match the `SeriesCollection(3)`

Excellent reference on the `SERIES` formula. [http://peltiertech.com/Excel/ChartsHowTo/ChartSeriesFormula.html](http://peltiertech.com/Excel/ChartsHowTo/ChartSeriesFormula.html)
