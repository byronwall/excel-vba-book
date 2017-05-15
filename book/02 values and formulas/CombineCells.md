## CombineCells.md

```vb
Public Sub CombineCells()

    'collect all user data up front
    Dim inputRange As Range
    On Error GoTo errHandler
    Set inputRange = GetInputOrSelection("Select the range of cells to combine")

    Dim delimiter As String
    delimiter = Application.InputBox("Delimeter:")
    If delimiter = "" Or delimiter = "False" Then GoTo delimiterError

    Dim outputRange As Range
    Set outputRange = GetInputOrSelection("Select the output range")
    
    'Check the size of input and adjust output
    Dim numberOfColumns As Long
    numberOfColumns = inputRange.Columns.Count
    
    Dim numberOfRows As Long
    numberOfRows = inputRange.Rows.Count
    
    outputRange = outputRange.Resize(numberOfRows, 1)
    
    'Read input rows into a single string
    Dim outputString As String
    Dim i As Long
    For i = 1 To numberOfRows
        outputString = vbNullString
        Dim j As Long
        For j = 1 To numberOfColumns
            outputString = outputString & delimiter & inputRange(i, j)
        Next
        'Get rid of the first character (delimiter)
        outputString = Right(outputString, Len(outputString) - 1)
        'Print it!
        outputRange(i, 1) = outputString
    Next
    Exit Sub
delimiterError:
    MsgBox "No Delmiter Selected!"
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
End Sub
```