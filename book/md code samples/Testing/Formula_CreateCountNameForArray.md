```vb
Public Sub Formula_CreateCountNameForArray()
    '---------------------------------------------------------------------------------------
    ' Procedure : Formula_CreateCountNameForArray
    ' Author    : @byronwall
    ' Date      : 2016 01 21
    ' Purpose   : meant to create formula with limited range of column
    '---------------------------------------------------------------------------------------
    '

    Dim rng_named As Range

    Dim str_name As String
    str_name = Application.InputBox("Name of the range", Type:=2)

    Set rng_named = ActiveWorkbook.Names(str_name).RefersToRange

    Dim str_form As String
    str_form = "=INDEX(" & str_name & ",1,1):INDEX(" & str_name & ",COUNTA(" & str_name & "),1)"

    ActiveWorkbook.Names.Add str_name & "_limited", str_form

End Sub
```