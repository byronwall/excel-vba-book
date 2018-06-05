### Keyboard Shortcuts

The simplest thing to do is to add keybaord shortcuts to your addin. There are two ways to do that:

* Open up the Macros form on the Developer tab. You can then hit "options" for a given Sub and assign a keyabord shortcut (TODO: add picture of htis)
* That approach can sometimes be a pain to edit later, so you can also add code to your addin ot add the shortcut.

The latter approach is nice because you can easily change the shortcut or the calling method. For addins, I will nearly always take the latter approach since it is much easier to deal with alter. For XLSM workbooks, I will do the former since it is easier to change from a workbook.

If you want to add the keyboard shortcut using code, use the code below.  Ideally, you would put this in a Workbook_Open event that is called when the workbook opens.  You can also use this approach to add/remove shortcuts depending on user input.

```vb
Public Sub SetUpKeyboardHooksForSelection()

    
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
