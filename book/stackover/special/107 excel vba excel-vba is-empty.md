# SO item 107
Code finds values from sheets and copies them over to one sheet. If a column is completely empty, it prints "NO ITEMS".

I need to make it so, once it is done copying the items over, it **finds any blank cells in column "B"** _(StartSht, "B")_ and from the range of the last occupied cell of "C" up, **fills it with the string "EMPTY"**

Any ideas how I would go about doing that?

It does (1) and I need it to do (2)

(1)

![enter image description here](https://i.stack.imgur.com/k4Lw4.png)

(2)

![enter image description here](https://i.stack.imgur.com/Cbj0z.png)

```
Set dict = GetValues(hc3.Offset(1, 0))
If dict.count > 0 Then                  
    'add the values to the master list, column 2
    Set d = StartSht.Cells(Rows.count, hc1.Column).End(xlUp).Offset(1, 0)
    d.Resize(dict.count, 1).Value = Application.Transpose(dict.items)
Else
    'if no items are under the HOLDER header
    StartSht.Range(StartSht.Cells(i, 2), StartSht.Cells(GetLastRowInColumn(StartSht, "C"), 1)) = " NO ITEMS "
End If

```

----

Blank cells are easy to find with the `SpecialCells` function. It is the same as using GoTo (or hitting <kbd>F5</kbd>) and choosing `Blanks`.

```
StartSheet.Range("B:B").SpecialCells(xlCellTypeBlanks).Value = "EMPTY"

```

You can do the same for column C after building the appropriate range.
