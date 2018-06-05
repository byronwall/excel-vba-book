## Chart_CopyToSheet.md

```vb
Public Sub Chart_CopyToSheet()

    Dim targetObject As ChartObject
    
    Dim selectedObject As Object
    Set selectedObject = Selection
    
    Dim newSheetResult As VbMsgBoxResult
    newSheetResult = MsgBox("Create a new sheet?", vbYesNo, "New sheet?")
    
    Dim targetSheet As Worksheet
    If newSheetResult = vbYes Then
        Set targetSheet = Worksheets.Add()
    Else: Set targetSheet = Application.InputBox("Pick a cell on an existing sheet", "Pick sheet", Type:=8).Parent
    End If
    
    For Each targetObject In Chart_GetObjectsFromObject(selectedObject)
        targetObject.Copy
        targetSheet.Paste
    Next targetObject
    
    targetSheet.Activate
End Sub
```