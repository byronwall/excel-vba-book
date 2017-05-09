```vb
Sub Rand_OpenFilesAndCopy()

    Dim sht_data As Worksheet
    Dim sht_output As Worksheet
    
    Set sht_output = ActiveSheet

    Dim path As Variant
    Dim folder As Variant
    
    Application.ScreenUpdating = False
    ' Another static folder
    folder = "O:\HCCShare\Operations\PE\Plant 8\Production Engineer\BWall\2013 11 Rheo troubleshooting\Recipes\PE7\2\"
    
    path = Dir(folder)
    
    Do While path <> ""

        Dim wkbk As Workbook
        Set wkbk = Workbooks.Open(folder & path)
        Set sht_data = wkbk.Sheets(1)
        sht_data.UsedRange.Copy
        
        sht_output.Cells(sht_output.UsedRange.Rows.Count + 1, 1) = wkbk.name
        sht_output.Cells(sht_output.UsedRange.Rows.Count, 2).PasteSpecial xlPasteValues
        
        wkbk.Close False
        
        path = Dir
    
    Loop

End Sub
```