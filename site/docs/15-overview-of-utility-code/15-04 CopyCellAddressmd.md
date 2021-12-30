## CopyCellAddress.md

```vb
Public Sub CopyCellAddress()


    'TODO: this need to get a button or a keyboard shortcut for easy use
    Dim myClipboard As MSForms.DataObject
    Set myClipboard = New MSForms.DataObject

    Dim sourceRange As Range
    Set sourceRange = Selection

    myClipboard.SetText sourceRange.Address(True, True, xlA1, True)
    myClipboard.PutInClipboard
End Sub
```
