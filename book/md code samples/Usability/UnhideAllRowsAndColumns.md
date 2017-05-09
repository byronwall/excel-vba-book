```vb
Public Sub UnhideAllRowsAndColumns()
    '---------------------------------------------------------------------------------------
    ' Procedure : UnhideAllRowsAndColumns
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Unhides everything in a Worksheet
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    ActiveSheet.Cells.EntireRow.Hidden = False
    ActiveSheet.Cells.EntireColumn.Hidden = False

End Sub
```