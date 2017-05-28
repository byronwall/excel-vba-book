## print layout and exporting

THis section will focus on the print and export related details of a Worksheet.  In particular, it will focus on the details that are typically accessed through the Page Layout menu.  This is one fo the unique asepcts of Worksheets becuase they are the holder of the print/export infomraiton.  The main detials related to this are:

* Print area
* Page layout -- this is a very large object with a lot of properties to be set
* Exporting and printing

The details in this section can be a real time saver because one of the more tedious asepcts of workign with Excel is ensuring that your reports/graphs/data will print or export correctly.  Being able to control these properties with VBA mkaes it possible to quickly apply the smae formatting to a large number of Worksheets without having to click nine million times.

TODO: add the content related to page layout

### Rand_common print settings

TODO: clean up this code

```vb
Sub Rand_CommonPrintSettings()

    Application.ScreenUpdating = False
    Dim sht As Worksheet

    For Each sht In Sheets
        sht.PageSetup.PrintArea = ""
        sht.ResetAllPageBreaks
        sht.PageSetup.PrintArea = ""
    
        With sht.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.75)
            .RightMargin = Application.InchesToPoints(0.75)
            .TopMargin = Application.InchesToPoints(1)
            .BottomMargin = Application.InchesToPoints(1)
            .HeaderMargin = Application.InchesToPoints(0.5)
            .FooterMargin = Application.InchesToPoints(0.5)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlLandscape
            .Draft = False
            .PaperSize = xlPaperLetter
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintErrors = xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = False
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End With
    Next sht
    
    Application.ScreenUpdating = True
End Sub
```
