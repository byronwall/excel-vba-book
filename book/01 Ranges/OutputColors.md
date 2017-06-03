## OutputColors.md

```vb
Public Sub OutputColors()

    Const MINIMUM_INTEGER As Long = 1
    Const MAXIMUM_INTEGER As Long = 10
    Dim i As Long
    For i = MINIMUM_INTEGER To MAXIMUM_INTEGER
        ActiveCell.Offset(i).Interior.Color = Chart_GetColor(i)
    Next i

End Sub
```