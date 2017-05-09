```vb
Public Sub UpdateScrollbars()
    '---------------------------------------------------------------------------------------
    ' Procedure : UpdateScrollbars
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Cheap trick that forces Excel to update the scroll bars after a large deletion
    '---------------------------------------------------------------------------------------
    '
    Dim targetRange As Variant
    targetRange = ActiveSheet.UsedRange.Address

End Sub
```