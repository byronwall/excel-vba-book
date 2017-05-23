# SO item 087
I'm trying to use STRCONV vbPropercase to correct input in a macro that I use to copy data from one "input" worksheet to another. I have been able to assign specific variables to cells and then apply the STRCONV Propercase to the variable, but no change is made to the text in the cell when I do this. The code saves the file as a combination of the variable names, so I know that it makes the corrections there. How do I make the change appear in the cell and not just to the variable in the code?

Here is an excerpt from the code I'm using:

```
Dim Property As String    
Dim Accnt As String

Property = Worksheets("Audit").Range("L6").Value
Property = StrConv(Property, vbProperCase)

Accnt = Worksheets("Audit").Range("L7").Value

ActiveWorkbook.saveas "D:\(username)\Documents\" & Accnt & " - " & Property & ".xlsx", FileFormat:= _
    xlOpenXMLWorkbook, CreateBackup:=False

```

----

Just flip the assignment around:

```
Worksheets("Audit").Range("L6").Value = Property

```
