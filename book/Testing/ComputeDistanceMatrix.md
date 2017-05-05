```vb
Public Sub ComputeDistanceMatrix()

    'get the range of inputs, along with input name
    Dim rng_input As Range
    Set rng_input = Application.InputBox("Select input data", "Input", Type:=8)

    'Dim rng_ID As Range
    'Set rng_ID = Application.InputBox("Select ID data", "ID", Type:=8)

    'turning off updates makes a huge difference here... could also use array for output
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    'create new workbook
    Dim wkbk As Workbook
    Set wkbk = Workbooks.Add

    Dim sht_out As Worksheet
    Set sht_out = wkbk.Sheets(1)
    sht_out.name = "scaled data"

    'copy data over to standardize
    rng_input.Copy wkbk.Sheets(1).Range("A1")

    'go to edge of data, add a column, add STANDARDIZE, copy paste values, delete

    Dim rng_data As Range
    Set rng_data = sht_out.Range("A1").CurrentRegion

    Dim rng_col As Range
    For Each rng_col In rng_data.Columns

        'edge cell
        Dim rng_edge As Range
        Set rng_edge = sht_out.Cells(1, sht_out.Columns.Count).End(xlToLeft).Offset(, 1)

        'do a normal dist standardization
        '=STANDARDIZE(A1,AVERAGE(A:A),STDEV.S(A:A))

        rng_edge.Formula = "=IFERROR(STANDARDIZE(" & rng_col.Cells(1, 1).Address(False, False) & ",AVERAGE(" & _
                           rng_col.Address & "),STDEV.S(" & rng_col.Address & ")),0)"

        'do a simple value over average to detect differences
        rng_edge.Formula = "=IFERROR(" & rng_col.Cells(1, 1).Address(False, False) & "/AVERAGE(" & _
                           rng_col.Address & "),1)"

        'fill that down
        Range(rng_edge, rng_edge.Offset(, -1).End(xlDown).Offset(, 1)).FillDown

    Next

    Application.Calculate
    sht_out.UsedRange.Value = sht_out.UsedRange.Value
    rng_data.EntireColumn.Delete

    Dim sht_dist As Worksheet
    Set sht_dist = wkbk.Worksheets.Add()
    sht_dist.name = "distances"

    Dim rng_out As Range
    Set rng_out = sht_dist.Range("A1")

    'loop through each row with each other row
    Dim rng_row1 As Range
    Dim rng_row2 As Range

    Set rng_input = sht_out.Range("A1").CurrentRegion

    For Each rng_row1 In rng_input.Rows
        For Each rng_row2 In rng_input.Rows

            'loop through each column and compute the distance
            Dim dbl_dist_sq As Double
            dbl_dist_sq = 0

            Dim int_col As Long
            For int_col = 1 To rng_row1.Cells.Count
                dbl_dist_sq = dbl_dist_sq + (rng_row1.Cells(1, int_col) - rng_row2.Cells(1, int_col)) ^ 2
            Next

            'take the sqrt of that value and output
            rng_out.Value = dbl_dist_sq ^ 0.5

            'get to next column for output
            Set rng_out = rng_out.Offset(, 1)
        Next

        'drop down a row and go back to left edge
        Set rng_out = rng_out.Offset(1).End(xlToLeft)
    Next

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    sht_dist.UsedRange.NumberFormat = "0.00"
    sht_dist.UsedRange.EntireColumn.AutoFit

    'do the coloring
    Formatting_AddCondFormat sht_dist.UsedRange

End Sub
```