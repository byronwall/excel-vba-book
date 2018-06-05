```vb
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
```