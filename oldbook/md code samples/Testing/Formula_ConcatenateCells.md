```vb
Public Sub Formula_ConcatenateCells()
    '---------------------------------------------------------------------------------------
    ' Procedure : Formula_ConcatenateCells
    ' Author    : @byronwall
    ' Date      : 2016 01 27
    ' Purpose   : will output a formula of concatenations based on cells
    '---------------------------------------------------------------------------------------
    '

    Dim rng_cell As Range
    Dim rng_joins As Range

    'get the cell to output to and the ranges to join
    Set rng_cell = GetInputOrSelection("Select the cell to put the formula")
    Set rng_joins = Application.InputBox("Select the cells to join", Type:=8)

    'get the separator
    Dim str_delim As String
    str_delim = Application.InputBox("What delimeter to use?")
    str_delim = "&""" & str_delim & """&"

    Dim arr_addr As Variant
    ReDim arr_addr(1 To rng_joins.Count)

    Dim int_count As Long
    int_count = 1

    Dim rng_join As Range
    For Each rng_join In rng_joins
        arr_addr(int_count) = rng_join.Address(False, False)
        int_count = int_count + 1
    Next

    Dim str_form As String
    str_form = "=" & Join(arr_addr, str_delim)

    rng_cell.Formula = str_form

End Sub
```