## GenerateRandomData.md

```vb
Public Sub GenerateRandomData()

    Const NUMBER_OF_ROWS As Long = 10
    Const NUMBER_OF_COLUMNS As Long = 3 '0 index
    Const DEFAULT_COLUMN_WIDTH As Long = 15
    
    'Since we only work with offset, targetcell can be a constant, but range constants are awkward
    Dim targetCell As Range
    Set targetCell = Range("B2")

    Dim i As Long

    For i = 0 To NUMBER_OF_COLUMNS
        targetCell.Offset(, i) = chr(65 + i)

        With targetCell.Offset(1, i).Resize(NUMBER_OF_ROWS)
            Select Case i
            Case 0
                .Formula = "=TODAY()+ROW()"
            Case Else
                .Formula = "=RANDBETWEEN(1,100)"
            End Select

            .Value = .Value
        End With
    Next i

    ActiveSheet.UsedRange.Columns.ColumnWidth = DEFAULT_COLUMN_WIDTH

End Sub
```