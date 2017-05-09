```vb
Public Function Chart_GetObjectsFromObject(ByVal inputObject As Object) As Variant
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GetObjectsFromObject
    ' Author    : @byronwall
    ' Date      : 2015 12 30
    ' Purpose   : Helper function which finds a valid ChartObject based on what is actually selected
    '             Returns a Collection (possibly empty) and should be handled with a For Each
    '---------------------------------------------------------------------------------------
    '
    Dim chartObjectCollection As New Collection

    'NOTE that this function does not work well with Axis objects.  Excel does not return the correct Parent for them.
    
    Dim targetObject As Variant
    Dim inputObjectType As String
    inputObjectType = TypeName(inputObject)

    Select Case inputObjectType
    
        Case "DrawingObjects"
            'this means that multiple charts are selected
            For Each targetObject In inputObject
                If TypeName(targetObject) = "ChartObject" Then
                    'add it to the set
                    chartObjectCollection.Add targetObject
                End If
            Next targetObject
            
        Case "Worksheet"
            For Each targetObject In inputObject.ChartObjects
                chartObjectCollection.Add targetObject
            Next targetObject
            
        Case "Chart"
            chartObjectCollection.Add inputObject.Parent
            
        Case "ChartArea", "PlotArea", "Legend", "ChartTitle"
            'parent is the chart, parent of that is the chart targetObject
            chartObjectCollection.Add inputObject.Parent.Parent
            
        Case "series"
            'need to go up three levels
            chartObjectCollection.Add inputObject.Parent.Parent.Parent
            
        Case "Axis", "Gridlines", "AxisTitle"
            'these are the oddly unsupported objects
            MsgBox "Axis/gridline selection not supported.  This is an Excel bug.  Select another element on the chart(s)."
    
        Case Else
            MsgBox "Select a part of the chart(s), except an axis."
    
    End Select

    Set Chart_GetObjectsFromObject = chartObjectCollection
End Function
```