## Colorize.md

The Sub below makes it possible to applied a banded row coloring based on changing values in the rows. That is, it looks at the value in a given cell and compares it to the cell above. If the value has changed, it applies the next color. If the same, it will apply the same color as the row above. This gives a simple demonstration of how it's possible to create simple or complicated formatting rules with VBA.

TODO: clean up the code below to simplify the process and show only the core bits needed.

```vb
Public Sub Colorize()

    Dim targetRange As Range
    On Error GoTo errHandler
    Set targetRange = GetInputOrSelection("Select range to color")
    Dim lastRow As Long
    lastRow = targetRange.Rows.Count
    Dim interiorColor As Long
    interiorColor = RGB(200, 200, 200)

    Dim sameColorForLikeValues As VbMsgBoxResult
    sameColorForLikeValues = MsgBox("Do you want to keep duplicate values the same color?", vbYesNo)

    If sameColorForLikeValues = vbNo Then

        Dim i As Long
        For i = 1 To lastRow
            If i Mod 2 = 0 Then
                targetRange.Rows(i).Interior.Color = interiorColor
            Else: targetRange.Rows(i).Interior.ColorIndex = xlNone
            End If
        Next
    End If


    If sameColorForLikeValues = vbYes Then
        Dim flipFlag As Boolean
        For i = 2 To lastRow
            If targetRange.Cells(i, 1) <> targetRange.Cells(i - 1, 1) Then flipFlag = Not flipFlag
            If flipFlag Then
                targetRange.Rows(i).Interior.Color = interiorColor
            Else: targetRange.Rows(i).Interior.ColorIndex = xlNone
            End If
        Next
    End If
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
End Sub
```
