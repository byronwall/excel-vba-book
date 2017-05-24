# SO item 110
I am working on my research and I have to generate more than 91 charts for each sheet and I would like to use macro to do that.

I am still new with macros but I tried to write this one, but it is not working. I would greatly appreciate your help on this issue!

The set of data that I have looks like this

```
> A1         B1     C1    D1    E1     F1    G1       H1    I1
> 
> 
> Period    Ratio       Period  Ratio       Period  ratio   
> 2000Q1    1.23        2000Q1  0.78        2000Q1  1.07    
> 2000Q2    1.43        2000Q2  1.12        2000Q2  0.76     2000Q3 1.8        
> 2000Q3    1.09        2000Q3  1.21

```

(under Columns A & B I have Period and Ratio) - then Column C is empty - then (under column D & E I have Period and Ratio) and so on.

I separated the data set with an empty column.

Please note that the are other rows (I have an update button that I every time I click a new row with (period- ratio) will be added for all the columns)- also the first row with values starts at row 3

I want to create a chart for each set of data (here 3 charts)

The macro I wrote is a follows:

```
Sub loopChart()
Dim mychart As Chart
Dim c As Integer
Sheets("analysis").Select

c = 1
While c <> 0 #I put this condition so that the code will know that I have no more data set

    Set mychart = Charts.Add
    mychart.SetSourceData Source:=Range(cells(3, c)).CurrentRegion, PlotBy:=xlColumns
    c = c + 3
Wend

For Each mychart In Sheets("class").ChartObjects
    mychart.ChartType = xlLineMarkers
Next mychart

End Sub

```

I am not too sure of what I am doing is correct, but I am facing a trouble with the range. Also I know that this macro will create a new chart-sheet.

how can I create all the charts on the "analysis" sheet next to the values?

I would greatly appreciate anyone's help!!

----

If you want all of the charts on the `analysis` `Worksheet`, you can change the `Location` when the chart is created. I also changed the sheet name in the second loop to match the sheet name. You can add this line of code to the first loop though; the second loop is not necessary. If you do that, be sure to set a new reference to `mychart` on the location line.

```
Sub loopChart()
Dim mychart As Chart
Dim c As Integer
Sheets("analysis").Select

c = 1
While c <> 0 #I put this condition so that the code will know that I have no more data set

    Set mychart = Charts.Add
    mychart.SetSourceData Source:=Range(cells(3, c)).CurrentRegion, PlotBy:=xlColumns
    'change location to sheet
    mychart.Location xlLocationAsObject, "analysis"
    c = c + 3
Wend

For Each mychart In Sheets("analysis").ChartObjects
    mychart.ChartType = xlLineMarkers
Next mychart

End Sub

```
