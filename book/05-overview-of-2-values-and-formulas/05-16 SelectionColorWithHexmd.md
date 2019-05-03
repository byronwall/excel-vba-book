## Selection_ColorWithHex.md

```vb
Public Sub Selection_ColorWithHex()

    Dim targetCell As Range
    Dim targetRange As Range
    On Error GoTo errHandler
    Set targetRange = GetInputOrSelection("Select the range of cells to color")
    For Each targetCell In targetRange
        targetCell.Interior.Color = RGB( _
                                    WorksheetFunction.Hex2Dec(Mid(targetCell.Value, 2, 2)), _
                                    WorksheetFunction.Hex2Dec(Mid(targetCell.Value, 4, 2)), _
                                    WorksheetFunction.Hex2Dec(Mid(targetCell.Value, 6, 2)))

    Next targetCell
    Exit Sub
errHandler:
    MsgBox "No selection made!"
End Sub
```
