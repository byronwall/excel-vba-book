```vb
Sub EvaluateArrayFormulaOnNewSheet()
    '---------------------------------------------------------------------------------------
    ' Procedure : EvaluateArrayFormulaOnNewSheet
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Wacky thing to force an array formula to return as an array
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    'cut cell with formula
    Dim StrAddress As String
    Dim rngStart As Range
    Set rngStart = Sheet1.Range("J2")
    StrAddress = rngStart.Address

    rngStart.Cut

    'create new sheet
    Dim sht As Worksheet
    Set sht = Worksheets.Add

    'paste cell onto sheet
    Dim rngArr As Range
    Set rngArr = sht.Range("A1")
    sht.Paste rngArr

    'expand array formula size.. resize to whatever size is needed
    rngArr.Resize(3).FormulaArray = rngArr.FormulaArray

    'get your result
    Dim VarArr As Variant
    VarArr = Application.Evaluate(rngArr.CurrentArray.Address)

    ''''do something with your result here... it is an array


    'shrink the formula back to one cell
    Dim strFormula As String
    strFormula = rngArr.FormulaArray

    rngArr.CurrentArray.ClearContents
    rngArr.FormulaArray = strFormula

    'cut and paste back to original spot
    rngArr.Cut

    Sheet1.Paste Sheet1.Range(StrAddress)

    Application.DisplayAlerts = False
    sht.Delete
    Application.DisplayAlerts = True

End Sub
```