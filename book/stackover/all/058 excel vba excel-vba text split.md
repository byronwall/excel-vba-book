# SO item 058
I've copied a column of data starting from D9 to D-whatever, the copied data have both decimal value and text. The data varies in each cell in column D

## Example

D9 : 1675.87 L/s D10 : 1555.87 L/s D11 : 1635.87 L/s This is my code i tried..

```
    Dim c As Collection, K As Long
     Set c = New Collection
     K = 9
     On Error Resume Next
     For Each r In Range("D9:D" & Cells(Rows.Count, "D").End(xlUp).Row)
     ary = Split(r.Text, ",")
     For Each a In ary
     c.Add a, CStr(a)
     If Err.Number = 0 Then
     Cells(K, "E").Value = a
     K = K + 1
     Else
     Err.Number = 0
     End If
     Next a
     Next r
     On Error GoTo 0

```

**I want to split the data so that it will be D6 1675.87 and E6 L/s OR remove the L/s completely.**

I know this is simple for most people but I'm relatively new at this so any help would be good. Thank you. You are much appreciated.

----

As noted, `Split` is the easy way to do this. If you know that you will always have a single space you can get all of the cells very quickly with

```
Sub SplitAndRewrite()

    Dim rng_start As Range
    Set rng_start = Range("D6")

    Dim rng_cell As Range
    For Each rng_cell In Range(rng_start, rng_start.End(xlDown))
        rng_cell.Resize(, 2) = Split(rng_cell, " ")
    Next
End Sub

```

Code works by iterating through a _contiguous_ (uses `End`) column of values and applying `Split` on the . It then takes the two values and pops them back on top of the cell using `Resize` to expand the output by one column.

`Split` returns an array so it can be quickly output back into the spreadsheet.
