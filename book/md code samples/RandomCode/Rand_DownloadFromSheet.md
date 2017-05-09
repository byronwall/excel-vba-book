```vb
Sub Rand_DownloadFromSheet()

    Dim rng_addr As Range
    
    Dim str_folder As Variant
    'Another static folder
    str_folder = InputBox("Provide a folder location for output")
    
    For Each rng_addr In Range("B2:B35")
    
        Download_File rng_addr, str_folder & rng_addr.Offset(, 1)
    
    Next rng_addr

End Sub
```