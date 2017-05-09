```vb
Public Sub MakeHyperlinks()
    '---------------------------------------------------------------------------------------
    ' Procedure : MakeHyperlinks
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Converts a set of cells to hyperlink to their targetCell value
    '---------------------------------------------------------------------------------------
    '
    '+Changed to inputbox
    On Error GoTo errHandler
    Dim targetRange As Range
    Set targetRange = GetInputOrSelection("Select the range of cells to convert to hyperlink")
    
    'TODO: choose a better variable name
    Dim targetCell As Range
    For Each targetCell In targetRange
        ActiveSheet.Hyperlinks.Add Anchor:=targetCell, Address:=targetCell
    Next targetCell
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
End Sub
```