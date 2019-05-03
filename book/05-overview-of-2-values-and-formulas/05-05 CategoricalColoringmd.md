## CategoricalColoring.md

```vb
Public Sub CategoricalColoring()


    '+Get User Input
    Dim targetRange As Range
    On Error GoTo errHandler
    Set targetRange = GetInputOrSelection("Select Range to Color")

    Dim coloredRange As Range
    Set coloredRange = GetInputOrSelection("Select Range with Colors")

    '+Do Magic
    Application.ScreenUpdating = False
    Dim targetCell As Range
    Dim foundRange As Variant

    For Each targetCell In targetRange
        foundRange = Application.Match(targetCell, coloredRange, 0)
        '+ Matches font style as well as interior color
        If IsNumeric(foundRange) Then
            targetCell.Font.FontStyle = coloredRange.Cells(foundRange).Font.FontStyle
            targetCell.Font.Color = coloredRange.Cells(foundRange).Font.Color
            '+Skip interior color if there is none
            If Not coloredRange.Cells(foundRange).Interior.ColorIndex = xlNone Then
                targetCell.Interior.Color = coloredRange.Cells(foundRange).Interior.Color
            End If
        End If
    Next targetCell
    '+ If no fill, restore gridlines
    targetRange.Borders.LineStyle = xlNone
    Application.ScreenUpdating = True
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
    Application.ScreenUpdating = True
End Sub
```
