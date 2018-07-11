## DeleteAllCharts.md

```vb
Public Sub DeleteAllCharts()

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