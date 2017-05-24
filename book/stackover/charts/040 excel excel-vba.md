# SO item 040
I'm trying to loop through a range below and get runtime error 1004\. The highlighted row is this one here:

```
ActiveChart.SeriesCollection(i).Values = Worksheets("Chart Help").Range(Cells(10 + j, 5), Cells(10 + j, 1006))

```

Can anyone tell me what's wrong?

```
If Worksheets("Chart Help").Cells(4, 9 + j) <> " " Then
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(i).Name = Worksheets("Chart Help").Cells(4, 9 + j)
    ActiveChart.SeriesCollection(i).XValues = Worksheets("Chart Help").Range("J5:J1006")
    ActiveChart.SeriesCollection(i).Values = Worksheets("Chart Help").Range(Cells(10 + j, 5), Cells(10 + j, 1006))
    ActiveChart.SeriesCollection(i).Select

    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With

    i = i + 1
End If

j = j + 1

```

----

I assume your loop has the proper code and you simply didn't paste it all.

Your 2 calls to `Cells` refer to the `ActiveSheet` and not the `Chart Help` worksheet like you intend. You will need to prefix `Cells` with `Worksheets("Chart Help").Cells` for it to not error.

Something like this:

```
ActiveChart.SeriesCollection(i).Values = Worksheets("Chart Help").Range(Worksheets("Chart Help").Cells(10 + j, 5), Worksheets("Chart Help").Cells(10 + j, 1006))

```

Ideally you would define a reference to that worksheet to clean up the code. You also do not have to prefix the `Range` with the worksheet in this case. Those two ideas combined give:

```
Dim sht_chart As Worksheet
Set sht_chart = Worksheets("Chart Help")
ActiveChart.SeriesCollection(i).Values = Range(sht_chart.Cells(10 + j, 5), sht_chart.Cells(10 + j, 1006))

```
