# SO item 071
I am trying to write a macro in Excel to calculate the standard deviation of same text in column A taking the values from column B and giving the results in column C:

![spreadsheet ](https://i.stack.imgur.com/0GASH.png)

I did it manually by putting the equation`=STDEV.S(A2;A3;A4;A16)`for "aaa". But I need to do this automatically because I am doing another calculation and procedures which are completing by macros. Here is my code:

```
Option Explicit
Sub Main()
    CollectArray "A", "D"
    DoSum "D", "E", "A", "B"
End Sub

' collect array from a specific column and print it to a new one without duplicates
' params:
'           fromColumn - this is the column you need to remove duplicates from
'           toColumn - this will reprint the array without the duplicates
Sub CollectArray(fromColumn As String, toColumn As String)

    ReDim arr(0) As String

    Dim i As Long
    For i = 1 To Range(fromColumn & Rows.Count).End(xlUp).Row
        arr(UBound(arr)) = Range(fromColumn & i)
        ReDim Preserve arr(UBound(arr) + 1)
    Next i
    ReDim Preserve arr(UBound(arr) - 1)
    RemoveDuplicate arr
    Range(toColumn & "1:" & toColumn & Range(toColumn & Rows.Count).End(xlUp).Row).ClearContents
    For i = LBound(arr) To UBound(arr)
        Range(toColumn & i + 1) = arr(i)
    Next i
End Sub

' sums up values from one column against the other column
' params:
'           fromColumn - this is the column with string to match against
'           toColumn - this is where the SUM will be printed to
'           originalColumn - this is the original column including duplicate
'           valueColumn - this is the column with the values to sum
Private Sub DoSum(fromColumn As String, toColumn As String, originalColumn As String, valueColumn As String)
    Range(toColumn & "1:" & toColumn & Range(toColumn & Rows.Count).End(xlUp).Row).ClearContents
    Dim i As Long
    For i = 1 To Range(fromColumn & Rows.Count).End(xlUp).Row
        Range(toColumn & i) = WorksheetFunction.SumIf(Range(originalColumn & ":" & originalColumn), Range(fromColumn & i), Range(valueColumn & ":" & valueColumn))
    Next i
End Sub

Private Sub RemoveDuplicate(ByRef StringArray() As String)
    Dim lowBound$, UpBound&, A&, B&, cur&, tempArray() As String
    If (Not StringArray) = True Then Exit Sub
    lowBound = LBound(StringArray): UpBound = UBound(StringArray)
    ReDim tempArray(lowBound To UpBound)
    cur = lowBound: tempArray(cur) = StringArray(lowBound)
    For A = lowBound + 1 To UpBound
        For B = lowBound To cur
            If LenB(tempArray(B)) = LenB(StringArray(A)) Then
                If InStrB(1, StringArray(A), tempArray(B), vbBinaryCompare) = 1 Then Exit For
            End If
        Next B
        If B > cur Then cur = B
        tempArray(cur) = StringArray(A)
    Next A
    ReDim Preserve tempArray(lowBound To cur): StringArray = tempArray
End Sub

```

It would be nice if someone could please give me an idea or solution. The above code is for calculating the summation of same text values. Is there any way to modify my code to calculate the standard deviation?

----

Here is a formula and VBA route that gives you the `STDEV.S` for each set of items.

**Picture** shows the various ranges and results. My input is the same as yours, but I accidentally sorted it at one point so they don't line up.

![enter image description here](https://i.stack.imgur.com/APJ2e.png)

Some notes

*   `ARRAY` is the actual answer you want. `NON-ARRAY` showing for later.
*   I included the PivotTable to test the accuracy of the method.
*   `VBA` is the same answer as `ARRAY` calculated as a UDF which could be used elsewhere in your VBA.

**Formula** in cell `D3` is an array formula entered with CTRL+SHIFT+ENTER. That same formula is in `E3` without the array entry. Both have been copied down to the end of the data.

```
=STDEV.S(IF(B3=$B$3:$B$21,$C$3:$C$21))

```

Since it seems you need a VBA version of this, you can use the same formula in VBA and just wrap it in `Application.Evaluate`. This is pretty much how @Jeeped gets an answer, converting the range to values which meet the criteria.

**VBA Code** uses `Evaluate` to process a formula string built from the ranges given as input.

```
Public Function STDEV_S_IF(rng_criteria As Range, rng_criterion As Range, rng_values As Range) As Variant

    Dim str_frm As String

    'formula to reproduce
    '=STDEV.S(IF(B3=$B$3:$B$21,$C$3:$C$21))

    str_frm = "STDEV.S(IF(" & _
        rng_criterion.Address & "=" & _
        rng_criteria.Address & "," & _
        rng_values.Address & "))"

    'if you have more than one sheet, be sure it evalutes in the right context
    'or add the sheet name to the references above
    'single sheet works fine with just Application.Evaluate

    'STDEV_S_IF = Application.Evaluate(str_frm)
    STDEV_S_IF = Sheets("Sheet2").Evaluate(str_frm)

End Function

```

The formula in `F3` is the VBA UDF of the same formula as above, it is entered as a normal formula (although entering as an array does not affect anything) and is copied down to the end.

```
=STDEV_S_IF($B$3:$B$21,B3,$C$3:$C$21)

```

It is worth noting that `.Evaluate` processes this correctly as an array formula. You can compare this against the `NON-ARRAY` column included in the output. I am not certain how Excel knows to treat it this way. There was previously [a fairly extended conversion about how `Evaluate` process array formulas and determines the output](http://stackoverflow.com/questions/30314570/vba-long-array-formula-via-application-evaluate). This is tangentially related to that conversation.

And for completeness, here is the test of the `Sub` side of things. I am running this code in a module with a sheet other than `Sheet2` active. This emphasizes the ability of using `Sheets("Sheets2").Evaluate` for a multi-sheet workbook since my `Range` call is technically misqualified. Console output is included.

```
Sub test()

    Debug.Print STDEV_S_IF(Range("B3:B21"), Range("B3"), Range("C3:C21"))
    'correctly returns  206.301357242263

End Sub

```
