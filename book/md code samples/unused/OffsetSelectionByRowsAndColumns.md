## OffsetSelectionByRowsAndColumns.md

```vb
Public Sub OffsetSelectionByRowsAndColumns(ByVal numberOfRows As Long, ByVal numberOfColumns As Long)

    If TypeOf Selection Is Range Then

        'this error should only get called if the new range is outside the sheet boundaries
        On Error GoTo OffsetSelectionByRowsAndColumns_Exit

        Selection.Offset(numberOfRows, numberOfColumns).Select

        On Error GoTo 0
    End If

OffsetSelectionByRowsAndColumns_Exit:

End Sub
```