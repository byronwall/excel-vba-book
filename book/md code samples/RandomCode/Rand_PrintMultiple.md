```vb
Sub Rand_PrintMultiple()

    'go through the tags, pick one, put it in place
    
    'print out a PDF to a file
    
    Application.ScreenUpdating = False
    'Another static folder
    Dim rng_tag As Range
    Dim str_path As String
    str_path = InputBox("Provide a folder for output location")
    
    For Each rng_tag In Range("TAGS[TAG]").SpecialCells(xlCellTypeVisible)
        
        Range("C1") = rng_tag
        
        Sheets("SUMMARY").ExportAsFixedFormat xlTypePDF, str_path & rng_tag & ".PDF", , , , , , False
        
        'code is used to get a summary
        'Dim sht As Worksheet
        'Set sht = Sheets("ALL TAGS")
        
        'sht.Range("A1").EntireRow.Insert
        'sht.Range("A1") = rng_tag
        
        'Range("I8:L8").Copy
        'sht.Range("B1").PasteSpecial xlPasteValues
        'sht.Range("F1").Value = str_path & rng_tag & ".PDF"
    
    Next rng_tag
    
    Application.ScreenUpdating = True

End Sub
```