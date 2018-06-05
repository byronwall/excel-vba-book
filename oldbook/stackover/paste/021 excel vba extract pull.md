# SO item 021
I am writing a macro to do the following:

Every time I open a workbook, pull data from a closed workbook on my computer and copy that data to a sheet titled "Availability" starting in cell A1.

Currently, all that happens is "TRUE" is put into cell A1 on the Availability sheet.

Please help.

```
Sub OpenWorkbookToPullData()

    Dim sht As Worksheet
    Dim lastRow As Long
    lastRow = ActiveSheet.UsedRange.Rows.Count
    Set sht = ThisWorkbook.Worksheets(Sheet1.Name)
    Dim path As String
    path = "C:\users\" & Environ$("username") & _
    "\desktop\RC Switch Project\Daily Automation _
    Availability Report.xlsx"

    Dim currentWb As Workbook
    Set currentWb = ThisWorkbook

    Dim openWb As Workbook
    Set openWb = Workbooks.Open(path)

    Dim openWs As Worksheet
    Set openWs = openWb.Sheets("Automation Data")

    currentWb.Sheets("Availability").Range("A1") _
    = openWs.Range("A5:K" & LastRow).Select
    openWb.Close (False)

End Sub

```

----

As @Greg mentioned, the `.Select` is not needed. Once that is removed though, you will have a new problem where the two ranges are not the same size. `Range("A1")` is only 1 cell while the other range will be at least 11\. Your current VBA will only overwrite the values in the Range called for, which is `A1` here.

To get around this there are two approaches which work well.

# Resize

`Resize` the left hand side so that it is the same size as the right hand side.

```
Sub OpenWorkbookToPullData()

    Dim sht As Worksheet
    Dim lastRow As Long
    lastRow = ActiveSheet.UsedRange.Rows.Count
    Set sht = ThisWorkbook.Worksheets(Sheet1.Name)
    Dim path As String
    path = "C:\users\" & Environ$("username") & _
    "\desktop\RC Switch Project\Daily Automation Availability Report.xlsx"

    Dim currentWb As Workbook
    Set currentWb = ThisWorkbook

    Dim openWb As Workbook
    Set openWb = Workbooks.Open(path)

    Dim openWs As Worksheet
    Set openWs = openWb.Sheets("Automation Data")

    Dim rng_data As Range
    Set rng_data = openWs.Range("A5:K" & lastRow)

    currentWb.Sheets("Availability").Range("A1").Resize( _
        rng_data.Rows.Count, rng_data.Columns.Count).Value = rng_data.Value

    openWb.Close (False)

End Sub

```

# Copy/PasteSpecial

Actually `Copy` and then `PasteSpecial`.

```
Sub OpenWorkbookToPullData()

    Dim sht As Worksheet
    Dim lastRow As Long
    lastRow = ActiveSheet.UsedRange.Rows.Count
    Set sht = ThisWorkbook.Worksheets(Sheet1.Name)
    Dim path As String
    path = "C:\users\" & Environ$("username") & _
    "\desktop\RC Switch Project\Daily Automation Availability Report.xlsx"

    Dim currentWb As Workbook
    Set currentWb = ThisWorkbook

    Dim openWb As Workbook
    Set openWb = Workbooks.Open(path)

    Dim openWs As Worksheet
    Set openWs = openWb.Sheets("Automation Data")

    Dim rng_data As Range
    Set rng_data = openWs.Range("A5:K" & lastRow)

    rng_data.Copy
    currentWb.Sheets("Availability").Range("A1").PasteSpecial xlPasteValues

    openWb.Close (False)

End Sub

```

Since it looks like you are going for values anyways, I would use the `Copy/PasteSpecial` route for clarity in the code.
