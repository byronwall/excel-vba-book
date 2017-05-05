```vb
Public Sub Chart_GoToYRange()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GoToYRange
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Selects the y values used for the series
    '---------------------------------------------------------------------------------------
    '

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