```vb
Public Sub SeriesSplitIntoBins()
    '---------------------------------------------------------------------------------------
    ' Procedure : SeriesSplitIntoBins
    ' Author    : @byronwall
    ' Date      : 2015 11 03
    ' Purpose   : Code will break a column of continuous data into bins for plotting
    '---------------------------------------------------------------------------------------
    '
    Const LESS_THAN_EQUAL_TO_GENERAL As String = "<= General"
    Const GREATER_THAN_GENERAL As String = "> General"
    On Error GoTo ErrorNoSelection

    Dim selectedRange As Range
    Set selectedRange = Application.InputBox("Select category range with heading", Type:=8)
    Set selectedRange = Intersect(selectedRange, selectedRange.Parent.UsedRange) _
                                 .SpecialCells(xlCellTypeVisible, xlLogical + _
                                  xlNumbers + xlTextValues)

    Dim valueRange As Range
    Set valueRange = Application.InputBox("Select values range with heading", Type:=8)
    Set valueRange = Intersect(valueRange, valueRange.Parent.UsedRange)

    ''need to prompt for max/min/bins
    Dim maximumValue As Double, minimumValue As Double, binValue As Long

    minimumValue = Application.InputBox("Minimum value.", "Min", _
                                        WorksheetFunction.Min(selectedRange), Type:=1)
                                   
    maximumValue = Application.InputBox("Maximum value.", "Max", _
                                        WorksheetFunction.Max(selectedRange), Type:=1)
                                   
    binValue = Application.InputBox("Number of groups.", "Bins", _
                                    WorksheetFunction.RoundDown(Math.Sqr(WorksheetFunction.Count(selectedRange)), _
                                    0), Type:=1)

    On Error GoTo 0

    'determine default value
    Dim defaultString As Variant
    defaultString = Application.InputBox("Enter the default value", "Default", "#N/A")

    'detect cancel and exit
    If StrPtr(defaultString) = 0 Then Exit Sub

    ''TODO prompt for output location

    valueRange.EntireColumn.Offset(, 1).Resize(, binValue + 2).Insert
    'head the columns with the values

    ''TODO add a For loop to go through the bins

    Dim targetBin As Long
    For targetBin = 0 To binValue
        valueRange.Cells(1).Offset(, targetBin + 1) = minimumValue + (maximumValue - _
                                                      minimumValue) * targetBin / binValue
    Next

    'add the last item
    valueRange.Cells(1).Offset(, binValue + 2).FormulaR1C1 = "=RC[-1]"

    'FIRST =IF($D2 <=V$1,$U2,#N/A)
    '=IF(RC4 <=R1C,RC21,#N/A)

    'MID =IF(AND($D2 <=W$1, $D2>V$1),$U2,#N/A)  '''W current, then left
    '=IF(AND(RC4 <=R1C, RC4>R1C[-1]),RC21,#N/A)

    'LAST =IF($D2>AA$1,$U2,#N/A)
    '=IF(RC4>R1C[-1],RC21,#N/A)

    ''TODO add number format to display header correctly (helps with charts)

    'put the formula in for each column
    '=IF(RC13=R1C,RC16,#N/A)
    Dim formulaHolder As Variant
    formulaHolder = "=IF(AND(RC" & selectedRange.Column & " <=R" & _
                    valueRange.Cells(1).Row & "C," & "RC" & selectedRange.Column & ">R" & _
                    valueRange.Cells(1).Row & "C[-1]" & ")" & ",RC" & valueRange.Column & "," & _
                    defaultString & ")"

    Dim firstFormula As Variant
    firstFormula = "=IF(AND(RC" & selectedRange.Column & " <=R" & _
                    valueRange.Cells(1).Row & "C)" & ",RC" & valueRange.Column & "," & defaultString _
                    & ")"

    Dim lastFormula As Variant
    lastFormula = "=IF(AND(RC" & selectedRange.Column & " >R" & _
                    valueRange.Cells(1).Row & "C)" & ",RC" & valueRange.Column & "," & defaultString _
                    & ")"

    Dim formulaRange As Range
    Set formulaRange = valueRange.Offset(1, 1).Resize(valueRange.Rows.Count - 1, binValue + 2)
    formulaRange.FormulaR1C1 = formulaHolder

    'override with first/last
    formulaRange.Columns(1).FormulaR1C1 = firstFormula
    formulaRange.Columns(formulaRange.Columns.Count).FormulaR1C1 = lastFormula

    formulaRange.EntireColumn.AutoFit

    'set the number formats

    formulaRange.Offset(-1).Rows(1).Resize(1, binValue + 1).NumberFormat = LESS_THAN_EQUAL_TO_GENERAL
    formulaRange.Offset(-1).Rows(1).Offset(, binValue + 1).NumberFormat = GREATER_THAN_GENERAL

    Exit Sub

ErrorNoSelection:
    'TODO: consider removing this prompt
    MsgBox "No selection made.  Exiting.", , "No selection"

End Sub
```