# SO item 083
I am struggling to try writing an excel 2013 vba code statement to Convert a table cell type (2,4) into a cell type "D2", I tried the following statement:

```
dim celdactiva as range

with mcard  **(this is an excel objectlist that contains cell(2,4))**

set celdactiva=.Cells(2, 4).Address(RowAbsolute:=False, ColumnAbsolute:=False)
to returns "D2". **But the part (RowAbsolute:=False , Pups up as error)**

end with

```

----

The error is really on the `Set` part. `Address` returns a `String`. If you want to store this `String` in `celdactiva` then you need to change the declaration type from `Range` to `String` and remove the `Set` down below.

```
Dim celdactiva as String

With mcard
   celdactiva=.Cells(2, 4).Address(RowAbsolute:=False, ColumnAbsolute:=False)
End With

```
