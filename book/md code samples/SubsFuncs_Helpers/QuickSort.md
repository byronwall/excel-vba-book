```vb
Public Sub QuickSort(ByVal arrayToSort As Variant, Optional ByVal lowBound As Variant, Optional ByVal highBound As Variant)
    '---------------------------------------------------------------------------------------
    ' Procedure : QuickSort
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sorting implementation for arrays
    ' Source    : http://stackoverflow.com/a/152325/4288101
    '             http://en.allexperts.com/q/Visual-Basic-1048/string-manipulation.htm
    '---------------------------------------------------------------------------------------
    '
    Dim sortingVariant As Variant
    Dim swapHolder As Variant
    Dim temporaryLowBound As Long
    Dim temporaryHighBound As Long

    If IsMissing(lowBound) Then lowBound = LBound(arrayToSort)
    If IsMissing(highBound) Then highBound = UBound(arrayToSort)

    temporaryLowBound = lowBound
    temporaryHighBound = highBound

    sortingVariant = arrayToSort((lowBound + highBound) \ 2)

    While (temporaryLowBound <= temporaryHighBound)

        While (UCase(arrayToSort(temporaryLowBound)) < UCase(sortingVariant) And temporaryLowBound < highBound)
            temporaryLowBound = temporaryLowBound + 1
        Wend

        While (UCase(sortingVariant) < UCase(arrayToSort(temporaryHighBound)) And temporaryHighBound > lowBound)
            temporaryHighBound = temporaryHighBound - 1
        Wend

        If (temporaryLowBound <= temporaryHighBound) Then
            swapHolder = arrayToSort(temporaryLowBound)
            arrayToSort(temporaryLowBound) = arrayToSort(temporaryHighBound)
            arrayToSort(temporaryHighBound) = swapHolder
            temporaryLowBound = temporaryLowBound + 1
            temporaryHighBound = temporaryHighBound - 1
        End If

    Wend

    If (lowBound < temporaryHighBound) Then QuickSort arrayToSort, lowBound, temporaryHighBound
    If (temporaryLowBound < highBound) Then QuickSort arrayToSort, temporaryLowBound, highBound

End Sub
```