```vb
Public Sub SplitIntoColumns()
    '---------------------------------------------------------------------------------------
    ' Procedure : SplitIntoColumns
    ' Author    : @byronwall, @RaymondWise
    ' Date      : 2015 07 24
    ' Purpose   : Splits a targetCell into columns next to it based on a delimeter
    '---------------------------------------------------------------------------------------
    '
    Dim inputRange As Range

    Set inputRange = GetInputOrSelection("Select the range of cells to split")

    Dim targetCell As Range

    Dim delimiter As String
    delimiter = Application.InputBox("What is the delimeter?", , ",", vbOKCancel)
    If delimiter = "" Or delimiter = "False" Then GoTo errHandler
    For Each targetCell In inputRange

        Dim targetCellParts As Variant
        targetCellParts = Split(targetCell, delimiter)

        Dim targetPart As Variant
        For Each targetPart In targetCellParts

            Set targetCell = targetCell.Offset(, 1)
            targetCell = targetPart

        Next targetPart

    Next targetCell
    Exit Sub
errHandler:
    MsgBox "No Delimiter Defined!"
End Sub
```