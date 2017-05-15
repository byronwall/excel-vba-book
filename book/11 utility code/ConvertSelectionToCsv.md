## ConvertSelectionToCsv.md

```vb
Public Sub ConvertSelectionToCsv()

    Dim sourceRange As Range
    Set sourceRange = GetInputOrSelection("Choose range for converting to CSV")

    If sourceRange Is Nothing Then Exit Sub

    Dim outputString As String

    Dim dataRow As Range
    For Each dataRow In sourceRange.Rows
        
        Dim dataArray As Variant
        dataArray = Application.Transpose(Application.Transpose(dataRow.Rows.Value2))
        
        'TODO:  improve this to use another Join instead of string concats
        outputString = outputString & Join(dataArray, ",") & vbCrLf

    Next dataRow

    Dim myClipboard As MSForms.DataObject
    Set myClipboard = New MSForms.DataObject

    myClipboard.SetText outputString
    myClipboard.PutInClipboard

End Sub
```