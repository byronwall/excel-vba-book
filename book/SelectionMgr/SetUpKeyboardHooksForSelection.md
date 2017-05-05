```vb
Public Sub SetUpKeyboardHooksForSelection()
    '---------------------------------------------------------------------------------------
    ' Procedure : SetUpKeyboardHooksForSelection
    ' Author    : @byronwall
    ' Date      : 2016 09 29
    ' Purpose   : Creates hotkey events for the selection events
    '---------------------------------------------------------------------------------------
    '
    
    'SHIFT =    +
    'CTRL =     ^
    'ALT =      %

    'set up the keys for the selection mover
    Application.OnKey "^%{RIGHT}", "SelectionOffsetRight"
    Application.OnKey "^%{LEFT}", "SelectionOffsetLeft"
    Application.OnKey "^%{UP}", "SelectionOffsetUp"
    Application.OnKey "^%{DOWN}", "SelectionOffsetDown"
    
    'set up the keys for the indent level
    Application.OnKey "+^%{RIGHT}", "Formatting_IncreaseIndentLevel"
    Application.OnKey "+^%{LEFT}", "Formatting_DecreaseIndentLevel"

End Sub
```