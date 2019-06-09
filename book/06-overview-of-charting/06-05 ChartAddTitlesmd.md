## Chart_AddTitles.md

```vb
Public Sub Chart_AddTitles()

    Dim targetObject As ChartObject
    Const X_AXIS_TITLE As String = "x axis"
    Const Y_AXIS_TITLE As String = "y axis"
    Const SECOND_Y_AXIS_TITLE As String = "2and y axis"
    Const CHART_TITLE As String = "chart"

    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        With targetObject.Chart
            If Not .Axes(xlCategory).HasTitle Then
                .Axes(xlCategory).HasTitle = True
                .Axes(xlCategory).AxisTitle.Text = X_AXIS_TITLE
            End If

            If Not .Axes(xlValue, xlPrimary).HasTitle Then
                .Axes(xlValue).HasTitle = True
                .Axes(xlValue).AxisTitle.Text = Y_AXIS_TITLE
            End If

            '2015 12 14, add support for 2and y axis
            If .Axes.Count = 3 Then
                If Not .Axes(xlValue, xlSecondary).HasTitle Then
                    .Axes(xlValue, xlSecondary).HasTitle = True
                    .Axes(xlValue, xlSecondary).AxisTitle.Text = SECOND_Y_AXIS_TITLE
                End If
            End If

            If Not .HasTitle Then
                .HasTitle = True
                .ChartTitle.Text = CHART_TITLE
            End If
        End With
    Next targetObject

End Sub
```
