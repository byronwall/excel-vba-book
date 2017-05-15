```vb
Public Sub CreateConditionalsForFormatting()
    '---------------------------------------------------------------------------------------
    ' Procedure : CreateConditionalsForFormatting
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates a set of conditional formats for order of magnitude numbers
    '---------------------------------------------------------------------------------------
    '
    On Error GoTo errHandler
    Dim inputRange As Range
    Set inputRange = GetInputOrSelection("Select the range of cells to convert")
    'add these in as powers of 3, starting at 1 = 10^0
    Const ARRAY_MARKERS As String = " ,k,M,B,T,Q"
    Dim arrMarkers As Variant
    arrMarkers = Split(ARRAY_MARKERS, ",")
    
    Dim i As Long
    For i = UBound(arrMarkers) To 0 Step -1

        With inputRange.FormatConditions.Add(xlCellValue, xlGreaterEqual, 10 ^ (3 * i))
            .NumberFormat = "0.0" & Application.WorksheetFunction.Rept(",", i) & " "" " & arrMarkers(i) & """"
        End With

    Next
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
End Sub
```