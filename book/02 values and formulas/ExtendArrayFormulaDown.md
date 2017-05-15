## ExtendArrayFormulaDown.md

```vb
Public Sub ExtendArrayFormulaDown()

    Dim startingRange As Range
    Dim targetArea As Range


    Application.ScreenUpdating = False

    Set startingRange = Selection

    For Each targetArea In startingRange.Areas
    
        Dim targetCell As Range
        For Each targetCell In targetArea.Cells

            If targetCell.HasArray Then

                Dim formulaString As String
                formulaString = targetCell.FormulaArray

                Dim startOfArray As Range
                Dim endOfArray As Range

                Set startOfArray = targetCell.CurrentArray.Cells(1, 1)
                Set endOfArray = startOfArray.Offset(0, -1).End(xlDown).Offset(0, 1)

                targetCell.CurrentArray.Formula = vbNullString

                Range(startOfArray, endOfArray).FormulaArray = formulaString

            End If

        Next targetCell
    Next targetArea


    'Find the range of the new array formula
    'Save current formula and clear it out
    'Apply the formula to the new range
    Application.ScreenUpdating = True
End Sub
```