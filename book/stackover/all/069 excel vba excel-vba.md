# SO item 069
Is it possible to work in Excel with some metric suffix notation: If I write 1000, the cell shows `1k`. if I write 1000000 the cell shows `1M`.

I made two functions to make a workaround but maybe there's a more suitable solution.

```
Function lecI(cadena) As Double
    u = Right(cadena, 1)
    If u = "k" Then
        mult = 1000
    ElseIf u = "M" Then
        mult = 1000000
    ElseIf u = "m" Then
        mult = 0.001
    End If
    lecI = Val(Left(cadena, Len(cadena) - 1)) * mult
End Function

Function wriI(num) As String
    If num > 1000000 Then 'M
        wriI = Str(Round(num / 1000000, 2)) & "M"
    ElseIf num > 1000 Then 'k
        wriI = Str(Round(num / 1000, 1)) & "k"
    ElseIf num < 0.01 Then 'm
        wriI = Str(Round(num * 1000, 1)) & "m"
    Else: wriI = Str(num)
    End If

```

----

Based on the [link by @Vasily](http://stackoverflow.com/questions/30544932/is-it-possible-to-use-metric-k-notations-in-excel/30577371#comment49162998_30544932), you can get the desired outcome using only Conditional Formatting. This is nice because it means that all of **your values are stored as Numbers and not Text and math works like normal**.

Overall steps:

*   Create a new conditional formatting for each block of 1000 that applies the number format for that block
*   Add the largest condition at the top so it formats first
*   Rinse and repeat to get all the ones you want

**Conditional formatting** used to style column `C` which is just random data at different powers of ten. It is the same number as column `D` just styled differently.

![conditional setup](https://i.stack.imgur.com/SLI9a.png)

**Number formats**, are pretty easy since they are the same as [that link, see Large Numbers section](http://www.excel-easy.com/examples/custom-number-format.html).

*   ones = `0 " "`
*   thousands = `0, " k"`
*   millions = `0,, " M"`
*   and so on for however many you want

**Automation**, if you don't want to click and type all day, here is some VBA that will create all the conditional formatting for you (for current `Selection`). This example goes out to billions. Keep adding powers of 3 by extending the `Array` with more entries.

```
Sub CreateConditionalsForFormatting()

    'add these in as powers of 3, starting at 1 = 10^0
    Dim arr_markers As Variant
    arr_markers = Array("", "k", "M", "B")

    For i = UBound(arr_markers) To 0 Step -1

        With Selection.FormatConditions.Add(xlCellValue, xlGreaterEqual, 10 ^ (3 * i))
            .NumberFormat = "0" & Application.WorksheetFunction.Rept(",", i) & " "" " & arr_markers(i) & """"
            .StopIfTrue = False
        End With

    Next

End Sub

```

I change the `StopIfTrue` value so that this does not break other conditional formatting that might exist. If the largest condition is at the top (added first) then the `NumberFormat` from that one holds. By default, these are created with `StopIfTrue = True`. This is a moot point if you do not have any other conditional formatting on these cells.
