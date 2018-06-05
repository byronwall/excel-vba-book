```vb
Sub AlphabetizeAndReportWithDupes()
    '''this one goes through a data source and alphabetizes it.
    '''keeping mainly for the select case and find/findnext
    Dim rng_data As Range
    Set rng_data = Range("B2:B28")

    Dim rng_output As Range
    Set rng_output = Range("I2")

    Dim arr As Variant
    arr = Application.Transpose(rng_data.Value)
    QuickSort arr
    'arr is now sorted

    Dim i As Long
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
```