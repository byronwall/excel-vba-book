# SO item 013
I an writing a UDF that needs to accept both Arrays and Ranges.

Usually declaring parameter as variant would work but a Range is an object so this no longer applies. That being said bellow I pasted code that only works when passing an array.
Here is theorethical example, based on `SUM`:

```
Function TSUM(numbers() As Variant) As Variant
    Dim i As Integer
    For i = 1 To UBound(numbers, 1)
        TSUM = TSUM + numbers(i)
    Next i
End Function

```

> =TSUM({1,1}) Returns 2
> =TSUM(A1:B1) Returns #VALUE!

So how can I fix above example to accept Ranges as well?

----

If you are content to sum the array/range item by item, I would just change to using a For Each loop that works well for either Ranges or Arrays.

Here is that version

```
Public Function TSUM(numbers As Variant) As Variant
    Dim i As Variant

    For Each i In numbers
        TSUM = TSUM + i
    Next i
End Function

```

If you generally want to work a function based on the type of the argument, you can use `TypeName()` and switching logic. Here is you function with that approach. I called it TSUM2 for uniqueness.

```
Public Function TSUM2(numbers As Variant) As Variant
    Dim i As Integer

    If TypeName(numbers) = "Range" Then
        TSUM2 = Application.WorksheetFunction.Sum(numbers)
    Else
        For i = 1 To UBound(numbers, 1)
            TSUM2 = TSUM2 + numbers(i)
        Next i
    End If
End Function

```

Note in both examples, I removed the parentheses from the numbers parameters (was `numbers() as Variant` before). This allows it to accept Range inputs.

If you take the second approach, be sure to debug and verify the TypeNames that could come through.
