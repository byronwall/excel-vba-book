# SO item 080
I'm just about finished writing this Sub for Excel. I'm basically asking my end user for a total (for example, `$3000`) find the total amount spent by each customer on the list and report those whose total is more than `$3000` (the amount provided by the user) on a new worksheet that I created called `Report`.

I have this code written so far, which also validates the value entered by the user:

```
Sub Userinput()

    Dim myValue As Variant
    myValue = InputBox("Give me some input")
    Range("E1").Value = myValue
    If (Len(myValue) < 0 Or Not IsNumeric(myValue)) Then
    MsgBox "Input not valid, code aborted.", vbCritical
    Exit Sub
    End If
End Sub

```

Any suggestions on how I can use the inputted value to search through the customer data base and find more than what was inputted and place that in a new worksheet?

EDIT: Data sample:

```
Customer orders         

Order Date  Customer ID Amount purchased    
02-Jan-12   190         $580    
02-Jan-12   144         $570    
03-Jan-12   120         $1,911  
03-Jan-12   192         $593    
03-Jan-12   145         $332    

```

----

Here is another approach which takes advantage of straight forward Excel features to `Copy` the customer IDs column, `RemoveDuplicates`, `SUMIF` based on customer, and `Delete` those rows over the minimum.

```
Sub CopyFilterAndCountIf()

    Dim dbl_min As Double
    dbl_min = InputBox("enter minimum search")

    Dim sht_data As Worksheet
    Dim sht_out As Worksheet

    Set sht_data = ActiveSheet
    Set sht_out = Worksheets.Add()

    sht_data.Range("B:B").Copy sht_out.Range("A:A")
    sht_out.Range("A:A").RemoveDuplicates 1, xlYes

    Dim i As Integer
    For i = sht_out.UsedRange.Rows.Count To 2 Step -1
        If WorksheetFunction.SumIf( _
            sht_data.Range("B:B"), sht_out.Cells(i, 1), sht_data.Range("C:C")) < dbl_min Then
            sht_out.Cells(i, 1).EntireRow.Delete
        End If
    Next
End Sub

```

I don't do error checking on the input, but you can add that in. I am also taking advantage of Excel's willingness to process entire columns instead of dealing with finding ranges. Definitely makes it easier to understand the code.

It should also be mentioned that you can accomplish all of these same features by using a Pivot Table with a filter on the `Sum` and no VBA.
