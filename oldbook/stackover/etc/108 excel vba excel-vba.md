# SO item 108
I'm trying to use the value in cell F2 as the maximum value of a range selection within a larger range of numbers.

I've gotten as far as the below, but am getting a compile error: variable not defined.

```
Option Explicit

Sub SelectByValue(Rng1 As Range, MinimunValue As Double, MaximumValue As Double)

    Dim MyRange As Range
    Dim Cell As Object

     'Check every cell in the range for matching criteria.
    For Each Cell In Rng1
        If Cell.Value >= MinimunValue And Cell.Value <= MaximumValue Then
            If MyRange Is Nothing Then
                Set MyRange = Range(Cell.Address)
            Else
                Set MyRange = Union(MyRange, Range(Cell.Address))
            End If
        End If
    Next
     'Select the new range of only matching criteria
    MyRange.Select

End Sub

Sub CallSelectByValue()

     'Call the macro and pass all the required variables to it.
     'In the line below, change the Range, Minimum Value, and Maximum Value as needed
    Call SelectByValue(Range("A2:A41"), 1, Range(F2).Value)

End Sub

```

----

[TessellatingHeckler has the better answer](http://stackoverflow.com/a/30928154/4288101), but it's worth noting that you can get away with not using the quotes if you use brackets instead of `Range`.

```
Call SelectByValue( [A2:A41], 1, [F2].Value)

```

This syntax is generally discouraged because the brackets lead to ambiguous results which can make a mess at run time.
