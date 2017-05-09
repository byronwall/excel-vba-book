```vb
Public Sub ChartPropMove()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartPropMove
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets the "move or size" setting for all charts
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        targetObject.Placement = xlFreeFloating
    Next targetObject

End Sub
```