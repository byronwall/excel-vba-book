## properties and methods on the Worksheet

This section will focus on the specific properties and functions that exist for a Worksheet.

TODO: add more content

### LockAllSheets.md

TODO: clean up this code

```vb
Public Sub LockAllSheets()

    Dim userPassword As Variant
    userPassword = Application.InputBox("Password to lock")

    If Not userPassword Then
        MsgBox "Cancelled."
    Else
        Application.ScreenUpdating = False

        'Changed to ActiveWorkbook so if add-in is not installed, it will target the active book rather than the xlam
        Dim targetSheet As Worksheet
        For Each targetSheet In ActiveWorkbook.Sheets
            On Error Resume Next
            targetSheet.Protect (userPassword)
        Next

        Application.ScreenUpdating = True
    End If

End Sub
```

### OutputSheets.md

TODO: clean up this code

```vb
Public Sub OutputSheets()

    Dim outputSheet As Worksheet
    Set outputSheet = Worksheets.Add(Before:=Worksheets(1))
    outputSheet.Activate

    Dim outputRange As Range
    Set outputRange = outputSheet.Range("B2")

    Dim targetRow As Long
    targetRow = 0

    Dim targetSheet As Worksheet
    For Each targetSheet In Worksheets

        If targetSheet.name <> outputSheet.name Then

            targetSheet.Hyperlinks.Add _
                outputRange.Offset(targetRow), "", _
                "'" & targetSheet.name & "'!A1", , _
                targetSheet.name
            targetRow = targetRow + 1

        End If
    Next targetSheet

End Sub
```

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
