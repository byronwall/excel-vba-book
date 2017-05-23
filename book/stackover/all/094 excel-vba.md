# SO item 094
I have a vba function which calls Application.Index. My problem is that some times it returns a value of an item which doesn't exist in the lookup range! I have an isError check, but it returns false. Out of about 150 rows, 30 of them are returning incorrect - the other 120 are returning correct.

If anyone can help, it would be greatly appreciated! Here is my code:

```
Function getQtyOnHand(skuRng As Range, tc As Range, skuCol As Range) As Long
     Dim index, retVal As Long, sku As String
     sku = skuRng.Value
     index = Application.Match(sku, skuCol)
     If IsError(index) Then
       retVal = 0
     Else
       retVal = Application.index(tc, index, 0)
     End If
     getQtyOnHand = retVal
End Function

```

For clarity, here is the information being sent to this function:

```
    Dim totalCol As Range, stockSkuCol As Range
    Set totalCol = wbStock.Worksheets("MAIN").Range("F:F")
    Set stockSkuCol = wbStock.Worksheets("MAIN").Range("A:A")

    getQtyOnHand(ws.Range("F2"), totalCol, stockSkuCol)

```

Some further testing.... here's a totally separate function showing the incorrect output:

```
 Sub testIndex()
     Dim wb1 As Workbook, wb2 As Workbook, ws1 As Worksheet, ws2 As Worksheet
     Set wb1 = Workbooks("Output.xlsm")
     Set wb2 = Workbooks("STOCK.xlsx")
     Set ws1 = wb1.Worksheets("StockList")
     Set ws2 = wb2.Worksheets("MAIN")

     Dim c1 As Range
     Set c1 = ws1.Range("D131")

     Dim ind
     ind = Application.WorksheetFunction.Match(c1.Value, ws2.Range("A:A"))

     Debug.Print (c1.Value & " was found in row " & ind & " whose value is " & ws2.Range("A" & ind))

 End Sub

```

The debug.print output is:

```
ZM-101 was found in row 100 whose value is YK21222L

```

!!???? (by the way, 100 is the last row in this document)

Thanks, Davey

----

Your call to `Match` does not specify a search type. You call it with only 2 parameters which defaults to a `1`. You want to explicitly call it with a `0` for exact matches. This is better understood by using the formula version in a normal spreadsheet to see the effect of that final parameter.

Fix:

```
index = Application.Match(sku, skuCol, 0)

```

Same story on your later call to `Application.WorksheetFunction.Match(..,..)`
