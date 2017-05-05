```vb
Public Sub CopyCellAddress()
    '---------------------------------------------------------------------------------------
    ' Procedure : CopyCellAddress
    ' Author    : @byronwall
    ' Date      : 2015 12 03
    ' Purpose   : Copies the current cell address to the myClipboard for paste use in a formula
    '---------------------------------------------------------------------------------------
    '

    'TODO: this need to get a button or a keyboard shortcut for easy use
    Dim myClipboard As MSForms.DataObject
    Set myClipboard = New MSForms.DataObject

    Dim sourceRange As Range
    Set sourceRange = Selection

    myClipboard.SetText sourceRange.Address(True, True, xlA1, True)
    myClipboard.PutInClipboard
End Sub
```