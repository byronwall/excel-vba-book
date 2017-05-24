# SO item 092
I have code that adds and that if an answer is provided based on a certain criteria it adds itself to a list. As I have been troubleshooting the rest of the program I have accrued a lot of answers that have been added. If I select the cells it blinks in and out and if I try to delete manually it gets stuck in a 'loop' and in order to do anything I have to crash excel.

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'sorts supervisor add list
Const myCol As String = "H"
Const N As Long = 2
Dim r As Long
If Not Intersect(Target, Range(myCol & ":" & myCol)) Is Nothing Then
r = Cells(Rows.Count, myCol).End(xlUp).Row
If r < N Then r = N
With Sheet2.Sort
.SortFields.Clear
.SortFields.add Key:=Range(myCol & N), SortOn:=xlSortOnValues, _ Order:=xlAscending, DataOption:=xlSortNormal
.SetRange Range(myCol & N & ":" & myCol & r)
.Header = xlNo
.MatchCase = False
.Orientation = xlTopToBottom
.SortMethod = xlPinYin
.Apply
End With
For x = r To N + 1 Step -1
If Cells(x, myCol).Value = Cells(x - 1, myCol).Value Then Cells(x, _ myCol).Delete shift:=xlUp
Next 
End If

```

The error appears to occur on

```
If Cells(x, myCol).Value = Cells(x - 1, myCol).Value Then Cells(x, _ myCol).Delete shift:=xlUp

```

also it has had problems deleting the duplicates.

----

It is generally recommended to disable events inside event processing code if you are likely to trigger the event you are processing.

`Delete` will cause the Selection to change which will cause this event to fire again. [See this excellent post on the topic](http://stackoverflow.com/questions/13860894/ms-excel-crashes-when-vba-code-runs/13861640#13861640).

To resolve, add `Application.EnableEvents = False` at the start and then `True` at the end.

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Application.EnableEvents = False
    '...your code here
    Application.EnableEvents = True

End Sub

```
