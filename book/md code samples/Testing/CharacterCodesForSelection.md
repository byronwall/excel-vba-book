```vb
Public Sub CharacterCodesForSelection()
    '---------------------------------------------------------------------------------------
    ' Procedure : CharacterCodesForSelection
    ' Author    : @byronwall
    ' Date      : 2016 01 27
    ' Purpose   : will output each character in the text
    '---------------------------------------------------------------------------------------
    '

    Dim letter As Variant

    Dim rng_val As Range
    Set rng_val = Selection

    Dim i As Long
    For i = 1 To Len(rng_val.Value)
        MsgBox Asc(Mid(rng_val.Value, i, 1))
    Next

End Sub
```