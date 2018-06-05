## Chart_GoToYRange.md

```vb
Public Sub Chart_GoToYRange()


    If TypeName(Selection) = "Series" Then
        Dim b As New bUTLChartSeries
        b.UpdateFromChartSeries Selection

        b.Values.Parent.Activate
        b.Values.Activate
    Else
        MsgBox "Select a series in order to use this."
    End If

End Sub
```