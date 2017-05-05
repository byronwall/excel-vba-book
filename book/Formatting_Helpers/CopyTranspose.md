```vb
Public Sub CopyTranspose()
    '---------------------------------------------------------------------------------------
    ' Procedure : CopyTranspose
    ' Author    : @byronwall, @RaymondWise
    ' Date      : 2015 07 31
    ' Purpose   : Takes a range of cells and does a copy/tranpose
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    'If user cancels a range input, we need to handle it when it occurs
    On Error GoTo errCancel
    Dim selectedRange As Range
    
    Set selectedRange = GetInputOrSelection("Select your range")

    Dim outputRange As Range
    'Need to handle the error of selecting more than one cell
    Set outputRange = GetInputOrSelection("Select the output corner")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim startingCornerCell As Range
    Set startingCornerCell = selectedRange.Cells(1, 1)

    Dim startingCellRow As Long
    startingCellRow = startingCornerCell.Row
    Dim startingCellColumn As Long
    startingCellColumn = startingCornerCell.Column

    Dim outputRow As Long
    Dim outputColumn As Long
    outputRow = outputRange.Row
    outputColumn = outputRange.Column

    Dim targetCell As Range
    
    'We check for the intersection to ensure we don't overwrite any of the original data
    'There's probably a better way to do this than For Each
    For Each targetCell In selectedRange
        If Not Intersect(selectedRange, Cells(outputRow + targetCell.Column - startingCellColumn, outputColumn + targetCell.Row - startingCellRow)) Is Nothing Then
            MsgBox "Your destination intersects with your data"
            Exit Sub
        End If
    Next targetCell

    For Each targetCell In selectedRange
        ActiveSheet.Cells(outputRow + targetCell.Column - startingCellColumn, outputColumn + targetCell.Row - startingCellRow).Formula = targetCell.Formula
    Next targetCell

errCancel:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
End Sub
```