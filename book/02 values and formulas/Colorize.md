```vb
Public Sub Colorize()
    '---------------------------------------------------------------------------------------
    ' Procedure : Colorize
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates an alternating color band based on targetCell values
    '---------------------------------------------------------------------------------------
    '
    Dim targetRange As Range
    On Error GoTo errHandler
    Set targetRange = GetInputOrSelection("Select range to color")
    Dim lastRow As Long
    lastRow = targetRange.Rows.Count
    Dim interiorColor As Long
    interiorColor = RGB(200, 200, 200)
    
    Dim sameColorForLikeValues As VbMsgBoxResult
    sameColorForLikeValues = MsgBox("Do you want to keep duplicate values the same color?", vbYesNo)

    If sameColorForLikeValues = vbNo Then
        
        Dim i As Long
        For i = 1 To lastRow
            If i Mod 2 = 0 Then
                targetRange.Rows(i).Interior.Color = interiorColor
            Else: targetRange.Rows(i).Interior.ColorIndex = xlNone
            End If
        Next
    End If


    If sameColorForLikeValues = vbYes Then
        Dim flipFlag As Boolean
        For i = 2 To lastRow
            If targetRange.Cells(i, 1) <> targetRange.Cells(i - 1, 1) Then flipFlag = Not flipFlag
            If flipFlag Then
                targetRange.Rows(i).Interior.Color = interiorColor
            Else: targetRange.Rows(i).Interior.ColorIndex = xlNone
            End If
        Next
    End If
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
End Sub
```