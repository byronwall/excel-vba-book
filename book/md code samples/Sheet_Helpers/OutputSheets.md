```vb
Public Sub OutputSheets()
    '---------------------------------------------------------------------------------------
    ' Procedure : OutputSheets
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates a new worksheet with a list and link to each sheet
    '---------------------------------------------------------------------------------------
    '
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