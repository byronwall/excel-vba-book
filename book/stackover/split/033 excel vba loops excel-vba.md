# SO item 033
I have an Excel sheet, and column `A` has a list of product names. Most of the products have multiple variations, such as:

```
A1: LDP2-sm
A2: LDP2-med
A3: LDP2-lg
A4: LDP3-sm
A5: LDP3-med
A6: LDP3-lg
A7: LDP3-xlg

```

Here, `LDP2` is 1 product with 3 variations and `LDP3` is 1 product with 4 variations

How do I loop through this list and find the start of a new variation? In the above example, I'd want to find `A1` and `A4`. Then, I want to insert a row above each.

Here is the code I have so far:

```
Dim rw As Long
Dim lr As Long
Dim cnt As Long
lr = 500
rw = 2
cnt = 1
Do
    If Range("A" & cnt).Value = *FIRST VARIATION OF A NEW PRODUCT*
        Rows(rw).Insert Shift:=xlDown
        cnt = cnt + 1
    Else
        cnt = cnt + 1
    End If
    rw = rw + 1
Loop While rw <> lr

```

What code do I need for _FIRST VARIATION OF A NEW PRODUCT_?

It needs to determine if the value of the cell starts with a different prefix than the cell above it.

I won't know what the product name starts with or how many variations the product has, but I do know that the first portion of the product name will change. I.e. `LDP2`, `LDP3`, `LDP4`, etc.

----

`Split` is the function you are looking for if you will always have a single `-` in the name. It will give you each part of the name and you can then compare to the previous row.

Here is code that works for your example.

```
Sub SplitProductName()

    Dim rng_cell As Range
    Dim str_prev As String

    For Each rng_cell In Range(Range("A1"), Range("A1").End(xlDown))

        Dim parts As Variant

        parts = Split(rng_cell, "-")

        'check that it is different
        If parts(0) <> str_prev Then
            rng_cell.EntireRow.Insert xlUp
        End If

        'assign previous for next row
        str_prev = parts(0)

    Next rng_cell
End Sub

```

**Before**

![before macro](https://i.stack.imgur.com/E75G3.png)

**After**

![after macro](https://i.stack.imgur.com/SDRDP.png)
