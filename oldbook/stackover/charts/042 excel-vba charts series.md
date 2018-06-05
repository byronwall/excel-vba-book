# SO item 042
this should be pretty basic for many participants here. Also, this is my first question here so I'm sorry if i make something wrong.

I am trying to unselect a specific value from charts. Unselecting the category from the chart is the problem. Basically, from all the series in that chart. if one of the series name(Value) is iqual "NotThisSeries" then unselect it.

```
Sub exmpl()

Dim MySeries As Variant

ActiveSheet.ChartObjects("Chart 3").Activate

For Each MySeries In ActiveChart.SeriesCollection

  Select Case MySeries.XValues

    Case MySeries.Name = "NotThisSeries"

      ActiveChart.SeriesCollection(MySeries.Name).IsFiltered = True

    Case Else

  End Select

Next MySeries

End Sub

```

----

The main issue with your code is that you are using a `Select Case` when you seem to just need a simple `If`.

The corrected code for that is

```
Sub exmpl()

    Dim MySeries As Series
    ActiveSheet.ChartObjects("Chart 3").Activate

    For Each MySeries In ActiveChart.SeriesCollection
      If MySeries.Name = "NotThisSeries" Then
          ActiveChart.SeriesCollection(MySeries.Name).IsFiltered = True
        End If
    Next MySeries

End Sub

```

If you want to use the `Select Case` to handle other names, here is the correct way to do that:

```
Sub exmpl()

    Dim MySeries As Series
    ActiveSheet.ChartObjects("Chart 3").Activate

    For Each MySeries In ActiveChart.SeriesCollection
        Select Case MySeries.Name
            Case "NotThisSeries"
              ActiveChart.SeriesCollection(MySeries.Name).IsFiltered = True
        End Select
    Next MySeries

End Sub

```

**Edit**, here is the corresponding code to hide a category instead of a series.

```
Sub exmpl()

    Dim MySeries As Series
    ActiveSheet.ChartObjects("Chart 3").Activate

    Dim i As Integer
    Dim cat As ChartCategory

    For i = 1 To ActiveChart.ChartGroups(1).FullCategoryCollection.Count

        Set cat = ActiveChart.ChartGroups(1).FullCategoryCollection(i)

        If cat.Name = "NotThisSeries" Then
            cat.IsFiltered = True
        End If
    Next

End Sub

```

Here is another SO question that helped with the second part. See comments to the question. [Set an excel chart filter with VBA](http://stackoverflow.com/questions/25896074/set-an-excel-chart-filter-with-vba)
