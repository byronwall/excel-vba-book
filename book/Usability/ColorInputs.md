```vb
Public Sub ColorInputs()
    '---------------------------------------------------------------------------------------
    ' Procedure : ColorInputs
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Finds cells with no value and colors them based on having a formula?
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim targetCell As Range
    Const FIRST_COLOR_ACCENT As String = "msoThemeColorAccent1"
    Const SECOND_COLOR_ACCENT As String = "msoThemeColorAccent2"
    'This is finding cells that aren't blank, but the description says it should be cells with no values..
    For Each targetCell In Selection
        If targetCell.Value <> "" Then
            If targetCell.HasFormula Then
                targetCell.Interior.ThemeColor = FIRST_COLOR_ACCENT
            Else
                targetCell.Interior.ThemeColor = SECOND_COLOR_ACCENT
            End If
        End If
    Next targetCell

End Sub
```