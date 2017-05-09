```vb
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
```