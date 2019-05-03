## CutPasteTranspose.md

```vb
Public Sub CutPasteTranspose()


    '########Still Needs to address Issue#23#############
    On Error GoTo errHandler
    Dim sourceRange As Range
    'TODO #Should use new inputbox function
    Set sourceRange = Selection

    Dim outputRange As Range
    Set outputRange = Application.InputBox("Select output corner", Type:=8)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim topLeftCell As Range
    Set topLeftCell = sourceRange.Cells(1, 1)

    Dim topRow As Long
    topRow = topLeftCell.Row
    Dim leftColumn As Long
    leftColumn = topLeftCell.Column

    Dim outputRow As Long
    Dim outputColumn As Long
    outputRow = outputRange.Row
    outputColumn = outputRange.Column

    outputRange.Activate

    'Check to not overwrite
    Dim targetCell As Range
    For Each targetCell In sourceRange
        If Not Intersect(sourceRange, Cells(outputRow + targetCell.Column - leftColumn, outputColumn + targetCell.Row - topRow)) Is Nothing Then
            MsgBox ("Your destination intersects with your data. Exiting.")
            GoTo errHandler
        End If
    Next

    'this can be better
    For Each targetCell In sourceRange
        targetCell.Cut
        ActiveSheet.Cells(outputRow + targetCell.Column - leftColumn, outputColumn + targetCell.Row - topRow).Activate
        ActiveSheet.Paste
    Next targetCell

errHandler:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

End Sub
```
