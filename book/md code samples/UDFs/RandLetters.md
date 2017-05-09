```vb
Public Function RandLetters(ByVal letterCount As Long) As String
    '---------------------------------------------------------------------------------------
    ' Procedure : RandLetters
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : UDF that generates a sequence of random letters
    '---------------------------------------------------------------------------------------
    '
    Dim letterIndex As Long
    
    Dim letters() As String
    ReDim letters(1 To letterCount)
    
    For letterIndex = 1 To letterCount
        letters(letterIndex) = chr(Int(Rnd() * 26 + 65))
    Next
    
    RandLetters = Join(letters(), "")
    
End Function
```