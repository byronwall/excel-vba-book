Attribute VB_Name = "RandomCode"
Option Explicit

Sub ExportFilesFromFolder()
    '---------------------------------------------------------------------------------------
    ' Procedure : ExportFilesFromFolder
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Goes through a folder and process all workbooks therein
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    '###Needs error handling
    'TODO: consider deleting this Sub since it is quite specific
    Application.ScreenUpdating = False

    Dim file As Variant
    Dim path As String
    path = InputBox("What path?")

    file = Dir(path)
    While (file <> "")

        Debug.Print path & file

        Dim fileName As String

        fileName = path & file

        Dim wbActive As Workbook
        Set wbActive = Workbooks.Open(fileName)

        Dim wsActive As Worksheet
        Set wsActive = wbActive.Sheets("Case Summary")

        With ActiveSheet.PageSetup
            .TopMargin = Application.InchesToPoints(0.4)
            .BottomMargin = Application.InchesToPoints(0.4)
        End With

        wsActive.ExportAsFixedFormat xlTypePDF, path & "PDFs\" & file & ".pdf"

        wbActive.Close False

        file = Dir
    Wend

    Application.ScreenUpdating = True

End Sub

Sub EvaluateArrayFormulaOnNewSheet()
    '---------------------------------------------------------------------------------------
    ' Procedure : EvaluateArrayFormulaOnNewSheet
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Wacky thing to force an array formula to return as an array
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    'cut cell with formula
    Dim StrAddress As String
    Dim rngStart As Range
    Set rngStart = Sheet1.Range("J2")
    StrAddress = rngStart.Address

    rngStart.Cut

    'create new sheet
    Dim sht As Worksheet
    Set sht = Worksheets.Add

    'paste cell onto sheet
    Dim rngArr As Range
    Set rngArr = sht.Range("A1")
    sht.Paste rngArr

    'expand array formula size.. resize to whatever size is needed
    rngArr.Resize(3).FormulaArray = rngArr.FormulaArray

    'get your result
    Dim VarArr As Variant
    VarArr = Application.Evaluate(rngArr.CurrentArray.Address)

    ''''do something with your result here... it is an array


    'shrink the formula back to one cell
    Dim strFormula As String
    strFormula = rngArr.FormulaArray

    rngArr.CurrentArray.ClearContents
    rngArr.FormulaArray = strFormula

    'cut and paste back to original spot
    rngArr.Cut

    Sheet1.Paste Sheet1.Range(StrAddress)

    Application.DisplayAlerts = False
    sht.Delete
    Application.DisplayAlerts = True

End Sub

Sub MakeSeveralBoxesWithNumbers()

    Dim shp As Shape
    Dim sht As Worksheet

    Dim rng_loc As Range
    Set rng_loc = Application.InputBox("select range", Type:=8)

    Set sht = ActiveSheet

    Dim int_counter As Long

    For int_counter = 1 To InputBox("How many?")

        Set shp = sht.Shapes.AddTextbox(msoShapeRectangle, rng_loc.left, _
                                        rng_loc.top + 20 * int_counter, 20, 20)

        shp.Title = int_counter

        shp.Fill.Visible = msoFalse
        shp.Line.Visible = msoFalse

        shp.TextFrame2.TextRange.Characters.Text = int_counter

        With shp.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
            .Solid
        End With

    Next

End Sub

Sub CreatePdfOfEachXlsxFileInFolder()
    
    'pick a folder
    Dim folderDialog As FileDialog
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    folderDialog.Show
    
    Dim folderPath As String
    folderPath = folderDialog.SelectedItems(1) & "\"
    
    'find all files in the folder
    Dim filePath As String
    filePath = Dir(folderPath & "*.xlsx")

    Do While filePath <> ""

        Dim wkbkFile As Workbook
        Set wkbkFile = Workbooks.Open(folderPath & filePath, , True)
        
        Dim sht As Worksheet
        
        For Each sht In wkbkFile.Worksheets
            sht.Range("A16").EntireRow.RowHeight = 15.75
            sht.Range("A17").EntireRow.RowHeight = 15.75
            sht.Range("A22").EntireRow.RowHeight = 15.75
            sht.Range("A23").EntireRow.RowHeight = 15.75
        Next

        wkbkFile.ExportAsFixedFormat xlTypePDF, folderPath & filePath & ".pdf"
        wkbkFile.Close False

        filePath = Dir
    Loop
End Sub

Sub AlphabetizeAndReportWithDupes()
    '''this one goes through a data source and alphabetizes it.
    '''keeping mainly for the select case and find/findnext
    Dim rng_data As Range
    Set rng_data = Range("B2:B28")

    Dim rng_output As Range
    Set rng_output = Range("I2")

    Dim arr As Variant
    arr = Application.Transpose(rng_data.Value)
    QuickSort arr
    'arr is now sorted

    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        
        'if duplicate, use FindNext, else just Find
        Dim rng_search As Range
        Select Case True
        Case i = LBound(arr), UCase(arr(i)) <> UCase(arr(i - 1))
            Set rng_search = rng_data.Find(arr(i))
        Case Else
            Set rng_search = rng_data.FindNext(rng_search)
        End Select

        ''''do your report stuff in here for each row
        'copy data over
        rng_output.Offset(i - 1).Resize(, 6).Value = rng_search.Resize(, 6).Value

    Next i
End Sub


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

Function DownloadFileAsString(ByVal vWebFile As String) As String
    Dim oXMLHTTP As Object, i As Long, vFF As Long, oResp() As Byte

    'You can also set a ref. to Microsoft XML, and Dim oXMLHTTP as MSXML2.XMLHTTP
    Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    oXMLHTTP.Open "GET", vWebFile, False         'Open socket to get the website
    oXMLHTTP.Send                                'send request

    'Wait for request to finish
    Do While oXMLHTTP.readyState <> 4
        DoEvents
    Loop

    DownloadFileAsString = oXMLHTTP.responseText 'Returns the results as a byte array

    'Clear memory
    Set oXMLHTTP = Nothing
End Function

Function Download_File(ByVal vWebFile As String, ByVal vLocalFile As String) As Boolean
    Dim oXMLHTTP As Object, i As Long, vFF As Long, oResp() As Byte

    'You can also set a ref. to Microsoft XML, and Dim oXMLHTTP as MSXML2.XMLHTTP
    Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    oXMLHTTP.Open "GET", vWebFile, False         'Open socket to get the website
    oXMLHTTP.Send                                'send request

    'Wait for request to finish
    Do While oXMLHTTP.readyState <> 4
        DoEvents
    Loop

    oResp = oXMLHTTP.responseBody                'Returns the results as a byte array

    'Create local file and save results to it
    Dim oStream As Object
    If oXMLHTTP.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write oXMLHTTP.responseBody
        oStream.SaveToFile vLocalFile, 2         ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    End If

    'Clear memory
    Set oXMLHTTP = Nothing
End Function

Sub Rand_DownloadFromSheet()

    Dim rng_addr As Range
    
    Dim str_folder As Variant
    'Another static folder
    str_folder = InputBox("Provide a folder location for output")
    
    For Each rng_addr In Range("B2:B35")
    
        Download_File rng_addr, str_folder & rng_addr.Offset(, 1)
    
    Next rng_addr

End Sub

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


Sub Rand_DumpTextFromAllSheets()

    Dim c As Range
    Dim s As Worksheet
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim main As Workbook
    Set main = ActiveWorkbook
    
    Dim w As Workbook
    Dim sw As Worksheet
    
    Set w = Application.Workbooks.Add
    Set sw = w.Sheets.Add
    
    Dim Row As Long
    Row = 0
    For Each s In main.Sheets
        For Each c In s.UsedRange.SpecialCells(xlCellTypeConstants)
            sw.Range("A1").Offset(Row) = c
            Row = Row + 1
        Next c
    Next s

End Sub


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

Sub Rand_CopyPasteValuesIntoNewSheet()

    Dim sht_new As Worksheet
    Dim sht_current As Worksheet
    
    Set sht_current = ActiveSheet
    
    Set sht_new = Worksheets.Add
    sht_current.UsedRange.Copy
    sht_new.PasteSpecial xlPasteValuesAndNumberFormats
    

End Sub

Sub Rand_ConvertToString()

    Dim cell As Range
    Dim sel As Range
    
    Set sel = Selection
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each cell In Intersect(sel, sel.Parent.UsedRange)
        If Not IsEmpty(cell.Value) And Not cell.HasFormula Then
            cell.Value = CStr(cell.Value)
        End If
    Next cell
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Sub Rand_KeepCellsWithText()

    Selection.SpecialCells(xlCellTypeConstants).Select

End Sub

Sub Rand_DeleteHiddenSheets()

    Dim sht As Worksheet
    
    Application.DisplayAlerts = False
    
    For Each sht In Worksheets
        If sht.Visible = xlSheetHidden Then
            sht.Delete
        End If
    Next sht
    
    Application.DisplayAlerts = True

End Sub
