## GetInputOrSelection.md

```vb
Public Function GetInputOrSelection(ByVal userPrompt As String) As Range

    Dim defaultString As String
    
    If TypeOf Selection Is Range Then
        defaultString = Selection.Address
    End If
    
    On Error GoTo ErrorNoSelection
    Set GetInputOrSelection = Application.InputBox(userPrompt, Type:=8, Default:=defaultString)
    
    Exit Function
    
ErrorNoSelection:
    Set GetInputOrSelection = Nothing
    
End Function
```