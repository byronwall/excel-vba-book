# SO item 076
I have a script I put together to import multiple text files that are semi colon delimited in to new workbooks for each file that is selected. It is about 99% functional aside from the fact it seems to paste all of the data from the selected text files in a column after the correctly imported columns in the new workbook. I'm not entirely sure what part of the code is causing this to be pasted in to that particular row. Below is the main part of the code. Can anyone identify where the issue might be?

Also I just want to say thanks to the community here. I have learned a lot by going through other posts on this forum.

```
FilesToOpen = Application.GetOpenFilename _
  (FileFilter:="Text Files (*.txt), *.txt", _
  MultiSelect:=True, Title:="Text Files to Open")

For i = LBound(FilesToOpen) To UBound(FilesToOpen)
    Set wkb = Workbooks.Open(FilesToOpen(i))
    Set wks = wkb.ActiveSheet
     With wks.QueryTables.Add(Connection:= _
        "TEXT;" & FilesToOpen(i), Destination:=Range("A1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .RefreshStyle = xlInsertDeleteCells
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
Next i

```

----

Your issue is the use of `Workbooks.Open` to get a new `Workbook`. You are really opening the text file in the same way as going to `File->Open`. If you want a _new_ `Workbook` to dump the data, create it explicitly using `Set wkb = Workbooks.Add()` instead of your current call to `Workbooks.Open`. You are seeing the file's data because you opened the file first.

**Full code**

```
FilesToOpen = Application.GetOpenFilename _
  (FileFilter:="Text Files (*.txt), *.txt", _
  MultiSelect:=True, Title:="Text Files to Open")

For i = LBound(FilesToOpen) To UBound(FilesToOpen)
    Set wkb = Workbooks.Add()
    Set wks = wkb.ActiveSheet
     With wks.QueryTables.Add(Connection:= _
        "TEXT;" & FilesToOpen(i), Destination:=Range("A1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .RefreshStyle = xlInsertDeleteCells
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
Next i

```
