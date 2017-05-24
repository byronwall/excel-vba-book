# SO item 003
I would like to chart the states of a machine over a period of time. For example it may be "running" for 2 hours and then "stopped" for 1 hour, and there may be several times each state occurs. Using a stacked bar chart I'd like to display the state and the amount of time it stays in the state. I'm finding that excel is assigning a new color and legend entry to each new state instance even if that state has already occurred.
How can I make same-named states within a chart have the same color (e.g. every time "running" is displayed it has the same color and a single legend entry)? Thanks

----

The state name is stored as the Series Name. There is a series for each stack in the Chart. It is possible to iterate through series and style them based on the Series Name. It is also possible to remove entries from the Legend using the LegendEntries object.

Combining these into a loop, you can update the series color if it matches a title and then remove the item from the Legend if it is not one of the first two Series. This assumes that "running" and "stopped" alternate at the start and are the entries to keep in the Legend. If this is not the case, you could do more logic to spot the entries to keep.

```
Sub style_chart()

    Dim cht As Chart
    Dim ser As Series

    'uses the active chart... assume it is selected
    Set cht = ActiveChart

    With cht
        'reset legend so that it matches series
        .HasLegend = False
        .HasLegend = True

        'iterate backwards to delete
        For i = .SeriesCollection.Count To 1 Step -1
            Set ser = .SeriesCollection(i)

            'set series colors based on name
            If ser.Name = "running" Then
                ser.Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
            ElseIf ser.Name = "stopped" Then
                ser.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
            End If

            'delete the legend entry if after first 2
            If i > 2 Then
                .Legend.LegendEntries(i).Delete
            End If
        Next i
    End With

End Sub

```

# Before

![before](https://i.stack.imgur.com/SQn83.png)

# After

![after](https://i.stack.imgur.com/iCoBV.png)
