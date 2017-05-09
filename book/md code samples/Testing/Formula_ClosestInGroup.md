```vb
Public Sub Formula_ClosestInGroup()
    '---------------------------------------------------------------------------------------
    ' Procedure : Formula_ClosestInGroup
    ' Author    : @byronwall
    ' Date      : 2016 01 27
    ' Purpose   : Adds a formula that puts a given cell into a group of values based on closest value
    '---------------------------------------------------------------------------------------
    '

    Dim rng_check As Range
    Dim rng_group As Range
    Dim rng_cell As Range

    Set rng_cell = GetInputOrSelection("Select the cell to put the formula")
    Set rng_check = Application.InputBox("Select the cell to find the group of", Type:=8)
    Set rng_group = Application.InputBox("Select the group the cell belongs to", Type:=8)

    Dim str_form As String

    str_form = "=INDEX(" & rng_group.Address(True, True, xlA1, True) & _
               ",MATCH(MIN(ABS(" & rng_group.Address(True, True, xlA1, True) & "-" & _
               rng_check.Address(False, False) & ")),ABS(" & rng_group.Address(True, True, xlA1, True) & "-" & rng_check.Address(False, False) & "),0))"

    rng_cell.FormulaArray = str_form

End Sub
```