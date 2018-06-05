## RandLetters.md

```vb
Public Function RandLetters(ByVal letterCount As Long) As String

    Dim letterIndex As Long
    
    Dim letters() As String
    ReDim letters(1 To letterCount)
    
    For letterIndex = 1 To letterCount
        letters(letterIndex) = chr(Int(Rnd() * 26 + 65))
    Next
    
    RandLetters = Join(letters(), "")
    
End Function
```