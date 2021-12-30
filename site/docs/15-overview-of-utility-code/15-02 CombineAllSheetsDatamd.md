## CombineAllSheetsData.md

```vb
Public Sub CombineAllSheetsData()

    'create the new wktk and sheet
    Dim targetWorkbook As Workbook
    Dim sourceWorkbook As Workbook

    Set sourceWorkbook = ActiveWorkbook
    Set targetWorkbook = Workbooks.Add

    Dim targetWorksheet As Worksheet
    Set targetWorksheet = targetWorkbook.Sheets.Add

    Dim isFirst As Boolean
    isFirst = True

    Dim targetRow As Long
    targetRow = 1

    Dim sourceWorksheet As Worksheet
    For Each sourceWorksheet In sourceWorkbook.Sheets
        If sourceWorksheet.name <> targetWorksheet.name Then

            sourceWorksheet.Unprotect

            'get the headers squared up
            If isFirst Then
                'copy over all headers
                sourceWorksheet.Rows(1).Copy targetWorksheet.Range("A1")
                isFirst = False

            Else
                'search for missing columns
                Dim headerRow As Range
                For Each headerRow In Intersect(sourceWorksheet.Rows(1), sourceWorksheet.UsedRange)

                    'check if it exists
                    Dim matchingHeader As Variant
                    matchingHeader = Application.Match(headerRow, targetWorksheet.Rows(1), 0)

                    'if not, add to header row
                    If IsError(matchingHeader) Then targetWorksheet.Range("A1").End(xlToRight).Offset(, 1) = headerRow
                Next headerRow
            End If

            'find the PnPID column for combo
            Dim pIDColumn As Long
            pIDColumn = Application.Match("PnPID", targetWorksheet.Rows(1), 0)

            'find the PnPID column for data
            Dim pIDData As Long
            pIDData = Application.Match("PnPID", sourceWorksheet.Rows(1), 0)

            'add the data, row by row
            Dim targetCell As Range
            For Each targetCell In sourceWorksheet.UsedRange.SpecialCells(xlCellTypeConstants)
                If targetCell.Row > 1 Then

                    'check if the PnPID exists in the combo sheet
                    Dim sourceRow As Variant
                    sourceRow = Application.Match( _
                               sourceWorksheet.Cells(targetCell.Row, pIDData), _
                               targetWorksheet.Columns(pIDColumn), _
                               0)

                    'add new row if it did not exist and id number
                    If IsError(sourceRow) Then
                        sourceRow = targetWorksheet.Columns(pIDColumn).Cells(targetWorksheet.Rows.Count, 1).End(xlUp).Offset(1).Row
                        targetWorksheet.Cells(sourceRow, pIDColumn) = sourceWorksheet.Cells(targetCell.Row, pIDData)
                    End If

                    'get column
                    Dim columnNumber As Long
                    columnNumber = Application.Match(sourceWorksheet.Cells(1, targetCell.Column), targetWorksheet.Rows(1), 0)

                    'update combo data
                    targetWorksheet.Cells(sourceRow, columnNumber) = targetCell

                End If
            Next targetCell
        End If
    Next sourceWorksheet
End Sub
```
