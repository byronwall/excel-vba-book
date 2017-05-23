# SO item 074
I want to loop through a range of cells alphabetically to create a report in alphabetical order. I dont want to sort the sheet as the original order is important.

```
Sub AlphaLoop()

'This is showing N and Z in uppercase, why?
For Each FirstLetter In Array(a, b, c, d, e, f, g, h, i, j, k, l, m, N, o, p, q, r, s, t, u, v, w, x, y, Z)
    For Each SecondLetter In Array(a, b, c, d, e, f, g, h, i, j, k, l, m, N, o, p, q, r, s, t, u, v, w, x, y, Z)
        For Each tCell In Range("I5:I" & Range("I20000").End(xlUp).Row)
            If Left(tCell, 2) = FirstLetter & SecondLetter Then
                'Do the report items here
        End If
        Next
    Next
Next

End Sub

```

Note that this code is untested, only sorts by the first 2 letters and is time consuming as it has to loop through the text 676 times. Is there a better way than this?

----

One option is to create an array of the values, quick sort the array, and then iterate the sorted array to create the report. This works even if there are duplicates in the source data (**edited**).

**Picture of ranges and results** shows the data in the left box and the sorted "report" on the right. My report is just copying the data from the original row. You could do whatever at this point. I added the coloring after the fact to show the correspondence.

![results of sorting](https://i.stack.imgur.com/OsCEI.png)

**Code** runs through the data index, sorts the values, and then runs through them again to output the data. It is using `Find/FindNext` to get the original item from the sorted array.

```
Sub AlphabetizeAndReportWithDupes()

    Dim rng_data As Range
    Set rng_data = Range("B2:B28")

    Dim rng_output As Range
    Set rng_output = Range("I2")

    Dim arr As Variant
    arr = Application.Transpose(rng_data.Value)
    QuickSort arr
    'arr is now sorted

    Dim i As Integer
    For i = LBound(arr) To UBound(arr)

        'if duplicate, use FindNext, else just Find
        Dim rng_search As Range
        Select Case True
            Case i = LBound(arr), UCase(arr(i)) <> UCase(arr(i - 1))
                Set rng_search = rng_data.Find(arr(i))
            Case Else
                Set rng_search = rng_data.FindNext(rng_search)
        End Select

        ''''do your report stuff in here for each row
        'copy data over
        rng_output.Offset(i - 1).Resize(, 6).Value = rng_search.Resize(, 6).Value

    Next i
End Sub

'from http://stackoverflow.com/a/152325/4288101
'modified to be case-insensitive and Optional params
Public Sub QuickSort(vArray As Variant, Optional inLow As Variant, Optional inHi As Variant)

    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long

    If IsMissing(inLow) Then
      inLow = LBound(vArray)
    End If

    If IsMissing(inHi) Then
      inHi = UBound(vArray)
    End If

    tmpLow = inLow
    tmpHi = inHi

    pivot = vArray((inLow + inHi) \ 2)

    While (tmpLow <= tmpHi)

       While (UCase(vArray(tmpLow)) < UCase(pivot) And tmpLow < inHi)
          tmpLow = tmpLow + 1
       Wend

       While (UCase(pivot) < UCase(vArray(tmpHi)) And tmpHi > inLow)
          tmpHi = tmpHi - 1
       Wend

       If (tmpLow <= tmpHi) Then
          tmpSwap = vArray(tmpLow)
          vArray(tmpLow) = vArray(tmpHi)
          vArray(tmpHi) = tmpSwap
          tmpLow = tmpLow + 1
          tmpHi = tmpHi - 1
       End If

    Wend

    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub

```

**Notes on the code:**

*   I have taken the Quick Sort code from this [previous answer](http://stackoverflow.com/a/152325/4288101) and added `UCase` to the comparisons for case-insensitive searching and made parameters `Optional` (and `Variant` for that to work).
*   The `Find/FindNext` part is going through the original data and locating the sorted items therein. If a duplicate is found (that is, if the current value matches the previous value) then it uses `FindNext` starting at the previously found entry.
*   My report generation is just taking the values from the data table. `rng_search` holds the `Range` of the current item in the original data source.
*   I am using `Application.Tranpose` to force `.Value` to be a `1-D` array instead of the multi-dim like normal. See [this answer for that usage](http://stackoverflow.com/a/7651439/4288101). Transpose the array again if you want to output into a column again.
*   The `Select Case` bit is just a hacky way of doing short-circuit evaluation in VBA. See [this previous answer](http://stackoverflow.com/a/3245183/4288101) about the usage of that.
