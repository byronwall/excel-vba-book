```vb
Public Sub OpenContainingFolder()
    '---------------------------------------------------------------------------------------
    ' Procedure : OpenContainingFolder
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Open the folder that contains the ActiveWorkbook
    '---------------------------------------------------------------------------------------
    '
    Dim targetWorkbook As Workbook
    Set targetWorkbook = ActiveWorkbook

    If targetWorkbook.path <> "" Then
        targetWorkbook.FollowHyperlink targetWorkbook.path
    Else
        MsgBox "Open file is not in a folder yet."
    End If

End Sub
```