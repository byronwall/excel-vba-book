```vb
Sub Rand_DeleteHiddenSheets()

    Dim sht As Worksheet
    
    Application.DisplayAlerts = False
    
    For Each sht In Worksheets
        If sht.Visible = xlSheetHidden Then
            sht.Delete
        End If
    Next sht
    
    Application.DisplayAlerts = True

End Sub
```