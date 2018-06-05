# SO item 026
I'm a newbie to Excel VBA and now I have a difficulty in dealing with the **Pivot Table**. The pivot table is used to **count the scores and scores occurrence**.

*   I've finished the average function within the pivot table.**(COMP)**
*   Now I would like to split the scores occurrence range. For instance, how many persons scores are in the range of scores >=80, or >=70, or >=60, or even <60 (failed),
*   and after this, I also need to count the **passed ratio**

Here are some conditions that must use the Pivot Table:

*   One manager has many leaders, and one leader have many members, the members take the exams and count their scores.
*   So the view of members' scores should bind with the manager and leader, to see each manager's data measure.

Below is my VBA code using Pivot Table to achieve my requirements:

```
Private Sub CommandButton1_Click()
Dim objTable As PivotTable, objField As PivotField
Dim ws As Worksheet
Dim wsPivot As Worksheet

ActiveWorkbook.Sheets("Sheet1").Select
Range("A1").Select
Set objTable = Sheet1.PivotTableWizard

' Specify row and column fields
Set objField = objTable.PivotFields("Creator's Tower Lead")
objField.Orientation = xlRowField
Set objField = objTable.PivotFields("Creator's Manager")
objField.Orientation = xlRowField
Set objField = objTable.PivotFields("Scores")
objField.Orientation = xlColumnField
'Set objField = objTable.PivotFields("Average Score")
'objField.Orientation = xlRowField
'Set objField = objTable.PivotFields("Score>=80")
'objField.Orientation = xlRowField
'Set objField = objTable.PivotFields("Score 70-79")
'objField.Orientation = xlRowField
'Set objField = objTable.PivotFields("Score 60-69")
'objField.Orientation = xlRowField
'Set objField = objTable.PivotFields("Score <60")
'objField.Orientation = xlRowField
'Set objField = objTable.PivotFields("Pass Rate (Score >60)")
'objField.Orientation = xlRowField

' Specify a data field with its summary
' function and format.
With objTable
    Set objField = objTable.PivotFields("Scores")
    objField.Orientation = xlDataField
    objField.Function = xlAverage
End With

With objTable
    Set objField = objTable.PivotFields("Scores")
    objField.Orientation = xlDataField
    objField.Function = xlCount
End With

'With objTable
'    Set objField = objTable.PivotFields("Scores")
'    objField.Orientation = xlDataField
'    Select Case objField
'        Case Scores > 80
'             objField.Function = xlCount
'        Case Scores > 70
'             objField.Function = xlCount
'        Case Scores > 60
'             objField.Function = xlCount
'        Case Else
'             objField.Function = xlCount
'    End Select
'End With End Sub

```

I really frustrated to get the scores range in pivot table. And some experts suggest me use secondary row(add a new column), but I tried and it still didn't work. Due to my stack overflow reputation is not greater than 10 points, I cannot upload my snapshots to make it clearer, hope my description could help you understand this question. Thanks in advance.

----

Excel provides grouping functionality that works well for most cases. For this type of example, you can specify the range to group by. To do this in VBA is as easy as getting a reference to one of the cells in field and calling `Group` on it.

Here is the code

```
Sub GroupPivot()

    Range("D3").Group Start:=60, End:=90, By:=10

End Sub

```

And here is an image with the normal menu and the result of the grouping call.

![results](https://i.stack.imgur.com/JXvXV.png)
