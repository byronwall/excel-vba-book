## OpenContainingFolder.md

```vb
Public Sub OpenContainingFolder()

    Dim targetWorkbook As Workbook
    Set targetWorkbook = ActiveWorkbook

    If targetWorkbook.path <> "" Then
        targetWorkbook.FollowHyperlink targetWorkbook.path
    Else
        MsgBox "Open file is not in a folder yet."
    End If

End Sub
```