## LockAllSheets.md

```vb
Public Sub LockAllSheets()

    Dim userPassword As Variant
    userPassword = Application.InputBox("Password to lock")

    If Not userPassword Then
        MsgBox "Cancelled."
    Else
        Application.ScreenUpdating = False

        'Changed to activeworkbook so if add-in is not installed, it will target the active book rather than the xlam
        Dim targetSheet As Worksheet
        For Each targetSheet In ActiveWorkbook.Sheets
            On Error Resume Next
            targetSheet.Protect (userPassword)
        Next

        Application.ScreenUpdating = True
    End If

End Sub
```