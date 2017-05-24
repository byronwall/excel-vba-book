# SO item 100
I have automated a proper record input into the table that I use as a database, and when the table is filtered the input don't work.

So I have code this to unfilter DataBase before every record input.

```
Public Sub UnFilter_DB()
Dim ActiveS As String, CurrScreenUpdate As Boolean

CurrScreenUpdate = Application.ScreenUpdating
Application.ScreenUpdating = False
ActiveS = ActiveSheet.Name

    Sheets("DB").Activate
    Sheets("DB").Range("A1").Activate
    Sheets("DB").ShowAllData
    DoEvents
    Sheets(ActiveS).Activate

Application.ScreenUpdating = CurrScreenUpdate
End Sub

```

But now, it stays stuck on `Sheets("DB").ShowAllData` saying :

> ShowAllData method of Worksheet Class failed

because the table is already unfiltered...

And I don't know if it is better to **use an error handler** like `On Error Resume Next` or how to **detect if there is a filter or none**.

Any pointers would be welcome!

----

If you use `Worksheet.AutoFilter.ShowAllData` instead of `Worksheet.ShowAllData` it will not throw the error when nothing is filtered.

This assumes that `Worksheet.AutoFilterMode = True` because otherwise you will get an error about `AutoFilter` not being an object.

```
Public Sub UnFilter_DB()
Dim ActiveS As String, CurrScreenUpdate As Boolean

CurrScreenUpdate = Application.ScreenUpdating
Application.ScreenUpdating = False
ActiveS = ActiveSheet.Name

    Sheets("DB").Activate
    Sheets("DB").Range("A1").Activate
    Sheets("DB").AutoFilter.ShowAllData
    DoEvents
    Sheets(ActiveS).Activate

Application.ScreenUpdating = CurrScreenUpdate
End Sub

```
