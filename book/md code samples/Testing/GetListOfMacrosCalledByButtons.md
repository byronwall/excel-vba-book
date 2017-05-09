```vb
Public Sub GetListOfMacrosCalledByButtons()
    '---------------------------------------------------------------------------------------
    ' Procedure : GetListOfMacrosCalledByButtons
    ' Author    : @byronwall
    ' Date      : 2016 01 28
    ' Purpose   : prints out a list of macros that are assigned to shapes
    '---------------------------------------------------------------------------------------
    '

    Dim sht As Worksheet
    Dim shp As Shape

    For Each sht In Worksheets
        For Each shp In sht.Shapes
            If shp.OnAction <> "" Then
                Debug.Print shp.OnAction
            End If
        Next
    Next
End Sub
```