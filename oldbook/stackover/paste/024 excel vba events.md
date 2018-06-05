# SO item 024
I'am trying to prevent users from pasting other things than values in the template I'm developing. I use a macro to always paste values in the worksheet (see below). When users switch to another workbook this macro should be disabled. The problem is that I get error 91 when activating another workbook.

'the macro in a module

```
Sub AlwaysPasteValues()
  Selection.PasteSpecial Paste:=xlPasteValues
End Sub

```

'the code in this workbook

```
Public wb As Workbook

Private Sub Workbook_Activate()
  Application.MacroOptions Macro:="AlwaysPasteValues", Description:="AlwaysPasteValues", ShortcutKey:="v"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Set wb = Nothing
End Sub

Private Sub Workbook_Deactivate()
  With wb
    .MacroOptions Macro:="AlwaysPasteValues", Description:="AlwaysPasteValues", ShortcutKey:=""
  End With
End Sub

Private Sub Workbook_Open()
  Set wb = ThisWorkbook
End Sub

```

----

You are changing `.MacroOptions` on `wb` which does not exist as a property. `MacroOptions` is for the `Application`. Use the same code as the `Activate` and you should be good.

```
Private Sub Workbook_Deactivate()

    Application.MacroOptions Macro:="AlwaysPasteValues", Description:="AlwaysPasteValues", ShortcutKey:=""

End Sub

```
