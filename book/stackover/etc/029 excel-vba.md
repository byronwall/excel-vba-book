# SO item 029
```
Sub test1()

Dim Str As String
Dim Search As String
Dim Status As String
Str = Cells(2, 5).Value
Search = FDSA!Cells(2, 5).Value
Status = FDSA!Cells(2, 10).Value

    If InStr(Search, Str) = True Then
                Status = "ok"
    Else
         End If

End Sub

```

I will be building up from this with loops. I want to check if what is in Cells(2,5) is contained in FDSA!Cells(2,5). If it is true then I would like to mark FDSA!Cells(2,10) as ok. I am getting an object required message. This is what I could come up with after looking at examples and tutorials. Let me know if you have questions

Only second time working on VBA. Thanks in advance, Alexis M.

----

Your syntax for referencing the worksheet is incorrect. That is probably throwing the error. You need to call to `Worksheets("FDSA")` and not use the `FDSA!` call like you have.

Also, you will have to set the cell value equal to `Status` for this to work. Just changing `Status` will not write it back into the workbook.

Also `InStr` returns the location of the match. If you want to know if there was a match, you need to check that the return is `>0`. This code should run and hopefully is closer to correct than your current code.

```
Sub test1()

Dim Str As String
Dim Search As String

Str = Cells(2, 5).Value
Search = Worksheets("FDSA").Cells(2, 5).Value

    If InStr(Search, Str) > 0 Then
        Worksheets("FDSA").Cells(2, 10).Value = "ok"
    End If

End Sub

```
