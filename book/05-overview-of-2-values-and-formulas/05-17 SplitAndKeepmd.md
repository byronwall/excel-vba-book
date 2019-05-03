## SplitAndKeep.md

```vb
Public Sub SplitAndKeep()

    On Error GoTo SplitAndKeep_Error

    Dim rangeToSplit As Range
    Set rangeToSplit = GetInputOrSelection("Select range to split")

    If rangeToSplit Is Nothing Then
        Exit Sub
    End If

    Dim delimiter As Variant
    delimiter = InputBox("What delimeter to split on?")
    'StrPtr is undocumented, perhaps add documentation or change function
    If StrPtr(delimiter) = 0 Then
        Exit Sub
    End If

    Dim itemToKeep As Variant
    'Perhaps inform user to input the sequence number of the item to keep
    itemToKeep = InputBox("Which item to keep? (This is 0-indexed)")

    If StrPtr(itemToKeep) = 0 Then
        Exit Sub
    End If

    Dim targetCell As Range
    For Each targetCell In Intersect(rangeToSplit, rangeToSplit.Parent.UsedRange)

        Dim delimitedCellParts As Variant
        delimitedCellParts = Split(targetCell, delimiter)

        If UBound(delimitedCellParts) >= itemToKeep Then
            targetCell.Value = delimitedCellParts(itemToKeep)
        End If

    Next targetCell

    On Error GoTo 0
    Exit Sub

SplitAndKeep_Error:
    MsgBox "Check that a valid Range is selected and that a number was entered for which item to keep."
End Sub
```
