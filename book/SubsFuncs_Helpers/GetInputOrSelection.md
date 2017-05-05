```vb
Public Function GetInputOrSelection(ByVal userPrompt As String) As Range
    '---------------------------------------------------------------------------------------
    ' Procedure : GetInputOrSelection
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Provides a single Function to get the Selection or Input with error handling
    '---------------------------------------------------------------------------------------
    '
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