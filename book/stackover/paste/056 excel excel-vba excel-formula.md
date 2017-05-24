# SO item 056
I have a table that is pulling thousands of rows of data from a very large sheet. Some of the columns in the table are getting their data from every 5th row on that large sheet. In order to speed up the process of creating the cell references, I used an OFFSET formula to grab a cell from every 5th row:

```
=OFFSET('Large Sheet'!B$2572,(ROW(1:1)-1)*5,,)
=OFFSET('Large Sheet'!B$2572,(ROW(2:2)-1)*5,,)
=OFFSET('Large Sheet'!B$2572,(ROW(3:3)-1)*5,,)
=OFFSET('Large Sheet'!B$2572,(ROW(4:4)-1)*5,,)
=OFFSET('Large Sheet'!B$2572,(ROW(5:5)-1)*5,,)
etc...

```

OFFSET can eat up resources during calculation of large tables though, and I'm looking for a way to speed up/simplify my formula. Is there any easy way to convert the OFFSET formula into just a simple cell reference like:

```
='Large Sheet'!B2572
='Large Sheet'!B2577
='Large Sheet'!B2582
='Large Sheet'!B2587
='Large Sheet'!B2592
etc...

```

I can't just paste values either. This needs to be an active reference, because the large sheet will change.

Thanks for your help.

----

If you want to take a VBA approach to this, you can generate the references very quickly using simple `For` loops.

Here is some _very_ crude code which can get you started. It uses hard-coded sheet names and variables. I am really just trying to show the `i*5` part.

```
Sub CreateReferences()

    For i = 0 To 12
        For j = 0 To 5
            Sheet2.Range("H1").Offset(i, j).Formula = _
                "=Sheet1!" & Sheet1.Range("A5").Offset(i * 5, j).Address
        Next
    Next

End Sub

```

It works by building a quick formula using the `Address` from a reference to a cell on `Sheet1`. The only key here is have one index count cells in the "summary" rows and multiply by 5 to get the reference to the "master" sheet. I am starting at `A5` just to match the results from `INDEX`.

**Results** show the formula input for `H1` and over. I am comparing to the `INDEX` results generated above.

![results](https://i.stack.imgur.com/fo7Bo.png)
