```vb
Public Sub Selection_ColorWithHex()
    '---------------------------------------------------------------------------------------
    ' Procedure : Selection_ColorWithHex
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Colors a targetCell based on the hex value stored in the targetCell
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
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