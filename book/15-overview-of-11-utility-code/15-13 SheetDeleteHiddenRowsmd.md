## Sheet_DeleteHiddenRows.md

```vb
Public Sub Sheet_DeleteHiddenRows()
    'These rows are unrecoverable
    Dim shouldDeleteHiddenRows As VbMsgBoxResult
    shouldDeleteHiddenRows = MsgBox("This will permanently delete hidden rows. They cannot be recovered. Are you sure?", vbYesNo)
    
    If Not shouldDeleteHiddenRows = vbYes Then Exit Sub
        
    Application.ScreenUpdating = False
    
    'collect a range to delete at end, using UNION-DELETE
    Dim rangeToDelete As Range
    
    Dim counter As Long
    counter = 0
    With ActiveSheet
        Dim rowIndex As Long
        For rowIndex = .UsedRange.Rows.Count To 1 Step -1
            If .Rows(rowIndex).Hidden Then
                If rangeToDelete Is Nothing Then
                    Set rangeToDelete = .Rows(rowIndex)
                Else
                    Set rangeToDelete = Union(rangeToDelete, .Rows(rowIndex))
                End If
                counter = counter + 1
            End If
        Next rowIndex
    End With
    
    rangeToDelete.Delete
    
    Application.ScreenUpdating = True
    
    MsgBox (counter & " rows were deleted")
End Sub
```