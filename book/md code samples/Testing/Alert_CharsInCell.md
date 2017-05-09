```vb
Public Sub Alert_CharsInCell()
    'will go through the current cell and report the ASCII keys (good for tracking down apparently blank cells)

    Dim str As String
    str = Selection.Value
    
    Dim chars As Variant
    ReDim chars(1 To Len(str))
    
    Dim index As Integer
    For index = 1 To Len(str)
        chars(index) = Asc(Mid(str, index, 1))
    Next

    MsgBox Join(Array(str, Join(chars, " ")), vbCrLf)

End Sub
```