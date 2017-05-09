```vb
Sub FormatChartsWithAxis()

    Dim chtObj As ChartObject
    For Each chtObj In ActiveSheet.ChartObjects
        
        Dim cht As Chart
        Dim ax As Axis
        
        Set cht = chtObj.Chart
        Set ax = cht.Axes(xlValue)
        
        ax.MinimumScale = 70
        ax.MaximumScale = 270
        
        cht.HasTitle = True
        
        Set ax = cht.Axes(xlCategory)
        ax.TickLabels.NumberFormat = "m/d HH:mm"
    
    Next

End Sub
```