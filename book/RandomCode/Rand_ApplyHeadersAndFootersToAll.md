```vb
Sub Rand_ApplyHeadersAndFootersToAll()

    Dim sht As Worksheet
    Dim sht_hdr As Worksheet
    
    Set sht_hdr = ActiveSheet
    
    For Each sht In Sheets
        sht.PageSetup.LeftHeader = sht_hdr.PageSetup.LeftHeader
        sht.PageSetup.CenterHeader = sht_hdr.PageSetup.CenterHeader
        sht.PageSetup.RightHeader = sht_hdr.PageSetup.RightHeader
        sht.PageSetup.LeftFooter = sht_hdr.PageSetup.LeftFooter
        sht.PageSetup.CenterFooter = sht_hdr.PageSetup.CenterFooter
        sht.PageSetup.RightFooter = sht_hdr.PageSetup.RightFooter
    Next sht

End Sub
```