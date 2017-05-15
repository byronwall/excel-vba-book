```vb
Public Sub TrimSelection()
    '---------------------------------------------------------------------------------------
    ' Procedure : TrimSelection
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Trims whitespace from a targetCell's value
    '---------------------------------------------------------------------------------------
    '
    Dim rangeToTrim As Range
    On Error GoTo errHandler
    Set rangeToTrim = GetInputOrSelection("Select the formulas you'd like to convert to static values")

    'disable calcs to speed up
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    'force to only consider used range
    Set rangeToTrim = Intersect(rangeToTrim, rangeToTrim.Parent.UsedRange)

    Dim targetCell As Range
    For Each targetCell In rangeToTrim
        
        'only change if needed
        Dim temporaryTrimHolder As Variant
        temporaryTrimHolder = Trim(targetCell.Value)
        
        'added support for char 160
        'TODO add more characters to remove
        temporaryTrimHolder = Replace(temporaryTrimHolder, chr(160), vbNullString)
        
        If temporaryTrimHolder <> targetCell.Value Then targetCell.Value = temporaryTrimHolder

    Next targetCell

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub
errHandler:
    MsgBox "No Delimiter Defined!"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub
```