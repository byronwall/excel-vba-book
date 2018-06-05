```vb
Sub Rand_Matrix()
    'Takes a table of values and flattens it.
    Dim rng_left As Range
    Dim rng_top As Range
    Dim rng_body As Range
        
    Set rng_left = Application.InputBox("Select left column", Type:=8)
    Set rng_top = Application.InputBox("Select top column", Type:=8)
    
    Dim int_left As Long, int_top As Long
    
    Set rng_body = Range(Cells(rng_left.Row, rng_top.Column), _
                         Cells(rng_left.Rows(rng_left.Rows.Count).Row, rng_top.Columns(rng_top.Columns.Count).Column))
                            
    Dim sht_out As Worksheet
    Set sht_out = Application.Worksheets.Add()
    
    Dim rng_cell As Range
    
    Dim int_row As Long
    int_row = 1
    
    For Each rng_cell In rng_body.SpecialCells(xlCellTypeConstants)
        sht_out.Range("A1").Offset(int_row) = rng_left.Cells(rng_cell.Row - rng_left.Row + 1, 1)
        sht_out.Range("B1").Offset(int_row) = rng_top.Cells(1, rng_cell.Column - rng_top.Column + 1)
        sht_out.Range("C1").Offset(int_row) = rng_cell
        
        int_row = int_row + 1
    Next rng_cell

End Sub
```