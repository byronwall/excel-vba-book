```vb
Sub Rand_PrintMultiplePvVsOp()

    'go through the tags, pick one, put it in place
    
    'print out a PDF to a file
    
    Application.ScreenUpdating = False
    'Another static folder
    Dim rng_tag As Range
    Dim str_path As String
    str_path = InputBox("Provide a folder for output location")
    
    For Each rng_tag In Range("tag_table[TAG]").SpecialCells(xlCellTypeVisible)
        
        Range("Charts!C3") = rng_tag
        
        Sheets("CHARTS").ExportAsFixedFormat xlTypePDF, str_path & rng_tag & "-" & rng_tag.Offset(, 5) & ".PDF", , , , , , False
    
    Next rng_tag
    
    Application.ScreenUpdating = True

End Sub
```