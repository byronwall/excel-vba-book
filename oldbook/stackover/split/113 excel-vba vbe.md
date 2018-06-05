# SO item 113
In response to [this](http://stackoverflow.com/q/30985795/4996248) question I thought it would be fun to write a VBE macro that would automatically replace lines which look like

```
DimAll a, b, c, d As Integer

```

by

```
Dim a As Integer, b As Integer, c As Integer, d As Integer

```

In my first draft I just want to modify a single selected line. After establishing the appropriate references to get to the VBE object model (see [http://www.cpearson.com/excel/vbe.aspx](http://www.cpearson.com/excel/vbe.aspx) ) and playing around a bit I came up with:

```
Function ExpandDim(codeLine As String) As String
    Dim fragments As Variant
    Dim i As Long, n As Long, myType As String
    Dim last As Variant
    Dim expanded As String

    If UCase(codeLine) Like "*DIMALL*AS*" Then
        codeLine = Replace(codeLine, "dimall", "Dim", , , vbTextCompare)
        fragments = Split(codeLine, ",")
        n = UBound(fragments)
        last = Split(Trim(fragments(n)))
        myType = last(UBound(last))
        For i = 0 To n - 1 'excludes last fragment
            expanded = expanded & IIf(i = 0, "", ",") & fragments(i) & " As " & myType
        Next i
        expanded = expanded & IIf(n > 0, ",", "") & fragments(n)
        ExpandDim = expanded
    Else
        ExpandDim = codeLine
    End If
End Function

Sub DimAll()
    Dim myVBE As VBE
    Dim startLine As Long, startCol As Long
    Dim endLine As Long, endCol As Long
    Dim myLine As String
    Set myVBE = Application.VBE
    myVBE.ActiveCodePane.GetSelection startLine, startCol, endLine, endCol
    myLine = myVBE.ActiveCodePane.CodeModule.Lines(startLine, 1)
    Debug.Print ExpandDim(myLine)
    myVBE.ActiveCodePane.CodeModule.ReplaceLine startLine, ExpandDim(myLine)
End Sub

```

In another code module I had:

```
Sub test()
    DimAll a, b, c, d As Integer
    Debug.Print TypeName(a)
    Debug.Print TypeName(b)
    Debug.Print TypeName(c)
    Debug.Print TypeName(d)
End Sub

```

This is the weird part. When I highlight the line which begins DimAll a, and invoke my awkwardly named sub DimAll, in the immediate window I see

```
Dim a As Integer, b As Integer, c As Integer, d As Integer

```

which is as expected, but in the code module itself the line is changed to

```
Dim a, b, c, d As Integer

```

DimAll has been replaced by Dim -- but the rest of the line is unmodified. I suspect that the commas are confusing the ReplaceLine method. Any ideas of how to fix this?

----

When I run with the debugger, `myLine` changes value between the two calls. The `DimAll` becomes `Dim` on the second time through.

This is because you are replacing the value of `codeLine` once you enter the main `If` conditional inside the `ExpandDim Function`.

Create a new variable in that function and you should be fine... or pass it `ByVal` and you're good:

```
Function ExpandDim(ByVal codeLine As String) As String

```
