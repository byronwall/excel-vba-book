```vb
Public Function Chart_GetColor(ByVal index As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GetColor
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Returns a list of colors for styling chart series
    '---------------------------------------------------------------------------------------
    '
    Dim colors(1 To 10) As Variant

    colors(1) = RGB(31, 120, 180)
    colors(2) = RGB(227, 26, 28)
    colors(3) = RGB(51, 160, 44)
    colors(4) = RGB(255, 127, 0)
    colors(5) = RGB(106, 61, 154)
    colors(6) = RGB(166, 206, 227)
    colors(7) = RGB(178, 223, 138)
    colors(8) = RGB(251, 154, 153)
    colors(9) = RGB(253, 191, 111)
    colors(10) = RGB(202, 178, 214)

    Chart_GetColor = colors(index)

End Function
```