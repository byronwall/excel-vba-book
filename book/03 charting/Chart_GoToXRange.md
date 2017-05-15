```vb
Public Sub Chart_GoToXRange()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GoToXRange
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Selects the x value range that is used for the series
    '---------------------------------------------------------------------------------------
    '

    If TypeName(Selection) = "Series" Then
        Dim b As New bUTLChartSeries
        b.UpdateFromChartSeries Selection

        b.XValues.Parent.Activate
        b.XValues.Activate
    Else
        MsgBox "Select a series in order to use this."
    End If

End Sub
```