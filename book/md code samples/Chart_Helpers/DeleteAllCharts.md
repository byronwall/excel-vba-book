```vb
Public Sub DeleteAllCharts()
    '---------------------------------------------------------------------------------------
    ' Procedure : DeleteAllCharts
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Helper Sub to delete all charts on ActiveSheet
    '---------------------------------------------------------------------------------------
    '
    If MsgBox("Delete all charts?", vbYesNo) = vbYes Then
        Application.ScreenUpdating = False

        Dim chartObjectIndex As Long
        For chartObjectIndex = ActiveSheet.ChartObjects.Count To 1 Step -1

            ActiveSheet.ChartObjects(chartObjectIndex).Delete

        Next chartObjectIndex

        Application.ScreenUpdating = True

    End If
End Sub
```