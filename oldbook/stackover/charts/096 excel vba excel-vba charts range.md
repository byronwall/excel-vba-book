# SO item 096
I have made a userform so a person can input a start date and an end date so a line graph will display the desired information. Currently I have everything working except the range update syntax.

I am saving the address of the start date's data as Ad and the address of the end date's address as Add (Both are strings).

I then try to set the range using these but I am doing something wrong. here is the code.

```
Dim CellX1 As Integer
Dim CellY1 As Integer
Dim CellX2 As Integer
Dim CellY2 As Integer
Dim Ad As String
Dim Add As String

Sheets("Data").Activate
Cells(CellY1, CellX1).Activate
Ad = ActiveCell.Address 'set start address

Cells(CellY2, CellX2).Activate
Add = ActiveCell.Address 'set end address

Sheets("Graph").Activate                                   

ActiveSheet.ChartObjects("Chart 1").Activate
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(1).Name = "=""A3"""

```

This is the lines of code that i cant get to work:

```
ActiveChart.SeriesCollection(1).Values = "=Data!$Ad:Add"
ActiveChart.SeriesCollection(1).XValues = "=Time!$E:$F"

```

----

Should be able to `Set` those as the `Range` version. You also will do better to assign the series variables to actual `Ranges` instead of addresses as strings. Really, you should just `Set` directly which is what I have below.

Full code should be something like:

```
Dim CellX1 As Integer
Dim CellY1 As Integer
Dim CellX2 As Integer
Dim CellY2 As Integer
Dim Ad As Range
Dim Add As Range

Set Ad = Sheets("Data").Cells(CellY1, CellX1) 'set start address    
Set Add = Sheets("Data").Cells(CellY2, CellX2) 'set end address

Sheets("Graph").Activate

ActiveSheet.ChartObjects("Chart 1").Activate
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(1).Name = Range("A3")
ActiveChart.SeriesCollection(1).Values = Range(Ad, Add)
ActiveChart.SeriesCollection(1).XValues = Worksheets("Time").Range("$E:$F")

```

Note that I changed the type of variable for `Ad` and `Add` to `Range`. This makes it easier to create a start/end `Range` for the chart.
