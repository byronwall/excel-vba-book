## ColorForUnique.md

```vb
Public Sub ColorForUnique()

    Dim dictKeysAndColors As New Scripting.Dictionary
    Dim dictColorsOnly As New Scripting.Dictionary

    Dim targetRange As Range

    On Error GoTo ColorForUnique_Error

    Set targetRange = GetInputOrSelection("Select column to color")
    Set targetRange = Intersect(targetRange, targetRange.Parent.UsedRange)

    'We can colorize the sorting column, or the entire row
    Dim shouldColorEntireRow As VbMsgBoxResult
    shouldColorEntireRow = MsgBox("Do you want to color the entire row?", vbYesNo)

    Application.ScreenUpdating = False

    Dim rowToColor As Range
    For Each rowToColor In targetRange.Rows

        'allow for a multi column key if intial range is multi-column
        'TODO: consider making this another prompt... might (?) want to color multi range based on single column key
        Dim keyString As String
        If rowToColor.Columns.Count > 1 Then
            keyString = Join(Application.Transpose(Application.Transpose(rowToColor.Value)), "||")
        Else
            keyString = rowToColor.Value
        End If

        'new value, need a color
        If Not dictKeysAndColors.Exists(keyString) Then
            Dim randomColor As Long
createNewColor:
            randomColor = RGB(Application.RandBetween(50, 255), _
                            Application.RandBetween(50, 255), Application.RandBetween(50, 255))
            If dictColorsOnly.Exists(randomColor) Then
                'ensure unique colors only
                GoTo createNewColor 'This is a sub-optimal way of performing this error check and loop
            End If

            dictKeysAndColors.Add keyString, randomColor
        End If

        If shouldColorEntireRow = vbYes Then
            rowToColor.EntireRow.Interior.Color = dictKeysAndColors(keyString)
        Else
            rowToColor.Interior.Color = dictKeysAndColors(keyString)
        End If
    Next rowToColor

    Application.ScreenUpdating = True

    On Error GoTo 0
    Exit Sub

ColorForUnique_Error:
    MsgBox "Select a valid range or fewer than 65650 unique entries."

End Sub
```
