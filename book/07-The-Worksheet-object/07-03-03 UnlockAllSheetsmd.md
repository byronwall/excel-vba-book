### UnlockAllSheets.md

TODO: clean up this code

```vb
Public Sub UnlockAllSheets()

    Dim userPassword As Variant
    userPassword = Application.InputBox("Password to unlock")

    Dim errorCount As Long
    errorCount = 0

    If Not userPassword Then
        MsgBox "Cancelled."
    Else
        Application.ScreenUpdating = False
        'Changed to ActiveWorkbook so if add-in is not installed, it will target the active book rather than the xlam
        Dim targetSheet As Worksheet
        For Each targetSheet In ActiveWorkbook.Sheets
            'Let's keep track of the errors to inform the user
            If Err.Number <> 0 Then errorCount = errorCount + 1
            Err.Clear
            On Error Resume Next
            targetSheet.Unprotect (userPassword)

        Next targetSheet
        If Err.Number <> 0 Then errorCount = errorCount + 1
        Application.ScreenUpdating = True
    End If
    If errorCount <> 0 Then
        MsgBox (errorCount & " sheets could not be unlocked due to bad password.")
    End If
End Sub
```
