```vb
Sub MakePdf2()

    Dim path As String
    path = "C:\Documents\TDA\2942\Data analysis of files\Bed Cycle analysis\PDF\"
    
    Dim index As Variant
    index = "recent"
 
        
    Sheet1.ExportAsFixedFormat xlTypePDF, path & index & ".pdf"


End Sub
```