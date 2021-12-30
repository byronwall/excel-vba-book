## print layout and exporting

This section will focus on the print and export related details of a Worksheet. In particular, it will focus on the details that are typically accessed through the Page Layout menu. This is one of the unique aspects of Worksheets because they are the holder of the print/export information. The main details related to this are:

- Print area
- Page layout -- this is a very large object with a lot of properties to be set
- Exporting and printing

The details in this section can be a real time saver because one of the more tedious aspects of working with Excel is ensuring that your reports/graphs/data will print or export correctly. Being able to control these properties with VBA makes it possible to quickly apply the same formatting to a large number of Worksheets without having to click nine million times.

When editing the Page Layout, you can change nearly everything. The one thing to be aware of is related to printers. There are a number of settings in the Worksheet that are internally tied to the default (or active) printer. This shows up if you are attempting to set the page size specifically. IF you always use the same printer or have coworkers who use the same printers, you amy not notice these issues. It becomes a serious problem when you ar etrying to make code work for multiple different printers that support or identify page sizes differently.

The best way to see what is savialable for page settings is to record a macro and change one thing. Excel is a bit aggressive at including all possible settings that could have changed. This is very nice if you want to grab some settings nad work them into your code.

There are a couple of other items to describe so you know what they are:

- Using `Zoom` and `FitToPages` to set the number of pages that the output will be included in (TODO: review)
- TODO: add others

Also, be aware that changing the print settings is a per Worksheet change. This amy be obvious since the prepares are off the Worksheet, but it is easy to forget this. The nice thing however s that you can just iterate your Worksheets and apply the same settings to all of them. This is one of the greatest time savers compared to changing properties in Excel (TODO: can these be changed with multi selection?).

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
