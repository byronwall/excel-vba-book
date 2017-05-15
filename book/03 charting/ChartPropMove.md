## ChartPropMove.md

```vb
Public Sub ChartPropMove()

    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        targetObject.Placement = xlFreeFloating
    Next targetObject

End Sub
```