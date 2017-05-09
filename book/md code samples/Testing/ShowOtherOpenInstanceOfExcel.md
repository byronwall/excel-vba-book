```vb
Sub ShowOtherOpenInstanceOfExcel()
    Dim oXLApp As Object

    'this will work if the previous instance was opened before the current one
    
    On Error Resume Next
    Set oXLApp = GetObject(, "Excel.Application")
    On Error GoTo 0

    oXLApp.Visible = True

    Set oXLApp = Nothing
End Sub
```