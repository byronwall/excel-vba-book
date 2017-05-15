## SeriesSplit.md

```vb
Public Sub SeriesSplit()

    On Error GoTo ErrorNoSelection

    Dim selectedRange As Range
    Set selectedRange = Application.InputBox("Select category range with heading", Type:=8)
    Set selectedRange = Intersect(selectedRange, selectedRange.Parent.UsedRange).SpecialCells(xlCellTypeVisible, xlLogical + xlNumbers + xlTextValues)

    Dim valueRange As Range
    Set valueRange = Application.InputBox("Select values range with heading", Type:=8)
    Set valueRange = Intersect(valueRange, valueRange.Parent.UsedRange)

    On Error GoTo 0

    'determine default value
    Dim defaultString As Variant
    defaultString = InputBox("Enter the default value", , "#N/A")
    'strptr is undocumented
    'detect cancel and exit
    If StrPtr(defaultString) = 0 Then
        Exit Sub
    End If

    Dim dictCategories As New Dictionary

    Dim categoryRange As Range
    For Each categoryRange In selectedRange
        'skip the header row
        If categoryRange.Address <> selectedRange.Cells(1).Address Then dictCategories(categoryRange.Value) = 1
    Next categoryRange

    valueRange.EntireColumn.Offset(, 1).Resize(, dictCategories.Count).Insert
    'head the columns with the values

    Dim valueCollection As Variant
    Dim counter As Long
    counter = 1
    For Each valueCollection In dictCategories
        valueRange.Cells(1).Offset(, counter) = valueCollection
        counter = counter + 1
    Next valueCollection

    'put the formula in for each column
    '=IF(RC13=R1C,RC16,#N/A)
    Dim formulaHolder As Variant
    formulaHolder = "=IF(RC" & selectedRange.Column & " =R" & _
                 valueRange.Cells(1).Row & "C,RC" & valueRange.Column & "," & defaultString & ")"

    Dim formulaRange As Range
    Set formulaRange = valueRange.Offset(1, 1).Resize(valueRange.Rows.Count - 1, dictCategories.Count)
    formulaRange.FormulaR1C1 = formulaHolder
    formulaRange.EntireColumn.AutoFit

    Exit Sub

ErrorNoSelection:
    'TODO: consider removing this prompt
    MsgBox "No selection made.  Exiting.", , "No selection"

End Sub
```