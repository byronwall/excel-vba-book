```vb
Public Sub LockAllSheets()
    '---------------------------------------------------------------------------------------
    ' Procedure : LockAllSheets
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Locks all sheets with the same password
    '---------------------------------------------------------------------------------------
    '
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