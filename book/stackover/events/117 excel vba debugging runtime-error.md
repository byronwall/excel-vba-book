# SO item 117
I'm trying to add the date and time of the last time a row was modified to a specific column of that row with the following VBA Script:

```
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    ThisRow = Target.Row
    If Target.Row > 1 Then Range("K" & ThisRow).Value = Now()
End Sub

```

But it keeps throwing the following error:

> Run-time error '-2147417848 (80010108)':
> 
> Method 'Value' of object 'Range' failed

Can anyone explain why this is happening?

----

You are creating an infinite loop by changing a value inside a `Worksheet_Change` event without disabling events first. When I do something similar, I get a range of errors from `Out of stack space` first to `Method Range failed...`.

Do this instead:

```
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False

    ThisRow = Target.Row
    If Target.Row > 1 Then Range("K" & ThisRow).Value = Now()

    Application.EnableEvents = True
End Sub

```

Related post: [MS Excel crashes when vba code runs](http://stackoverflow.com/questions/13860894/ms-excel-crashes-when-vba-code-runs/13861640#13861640)
