```vb
Public Sub CopyDiscontinuousRangeValuesToClipboard()

    Dim rngCSV As Range
    Set rngCSV = GetInputOrSelection("Choose range for converting to CSV")

    If rngCSV Is Nothing Then
        Exit Sub
    End If

    'get the counts for rows/columns
    Dim int_row As Long
    Dim int_cols As Long

    Set rngCSV = Intersect(rngCSV, rngCSV.Parent.UsedRange)

    'build the string array
    Dim arr_rows() As String
    ReDim arr_rows(1 To rngCSV.Areas(1).Rows.Count) As String

    Dim bool_firstArea As Boolean
    bool_firstArea = True

    Dim rng_area As Range
    For Each rng_area In rngCSV.Areas
        For int_row = 1 To UBound(arr_rows)
            If bool_firstArea Then
                arr_rows(int_row) = Join(Application.Transpose(Application.Transpose(rng_area.Rows(int_row).Value)), vbTab)
            Else
                arr_rows(int_row) = arr_rows(int_row) & vbTab & Join(Application.Transpose(Application.Transpose(rng_area.Rows(int_row).Value)), vbTab)
            End If
        Next

        bool_firstArea = False
    Next

    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject

    clipboard.SetText Join(arr_rows, vbCrLf)
    clipboard.PutInClipboard

End Sub
```