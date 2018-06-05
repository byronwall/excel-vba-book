```vb
Sub PadWithSpaces()

    'quick and dirty function to add a bunch of spaces to the end of the ActiveCell

    Dim lng_spaces As Long
    lng_spaces = InputBox("How many spaces?")
    
    ActiveCell.Value = ActiveCell.Value & WorksheetFunction.Rept(" ", lng_spaces)

End Sub
```