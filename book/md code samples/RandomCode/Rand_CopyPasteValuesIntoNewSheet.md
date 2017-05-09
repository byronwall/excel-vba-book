```vb
Sub Rand_CopyPasteValuesIntoNewSheet()

    Dim sht_new As Worksheet
    Dim sht_current As Worksheet
    
    Set sht_current = ActiveSheet
    
    Set sht_new = Worksheets.Add
    sht_current.UsedRange.Copy
    sht_new.PasteSpecial xlPasteValuesAndNumberFormats
    

End Sub
```