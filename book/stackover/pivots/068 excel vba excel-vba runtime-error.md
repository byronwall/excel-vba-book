# SO item 068
I need help with this macro. Every time I run it, I get the error below. I thought it was a simple macro that I could have anybody on my team use to make it take a less time than they were taking to manually create this PivotTable every time they ran the report. However, it's not working. Please see error below and advise. I emboldened and italicized the error.

![Error](https://i.stack.imgur.com/cqmPhm.jpg)

```
Sub LEDOTTR()
'
' LEDOTTR Macro
'

'
    Range("A87").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ***ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R87C1:R8214C25", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="LED OTTR!R1C1", TableName:="PivotTable6", _
        DefaultVersion:=xlPivotTableVersion14***
    Sheets("LED OTTR").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("LED")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Hierarchy name")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable6").PivotFields("LED").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("LED")
        .PivotItems("LED Marine").Visible = False
        .PivotItems("LL48 Linear LED").Visible = False
        .PivotItems("Other").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable6").PivotFields("LED"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
        "PivotTable6").PivotFields("   Late " & Chr(10) & "Indicator"), "Sum of    Late " & Chr(10) & "Indicator", _
        xlSum
    ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
        "PivotTable6").PivotFields("Early /Ontime" & Chr(10) & "   Indicator"), _
        "Sum of Early /Ontime" & Chr(10) & "   Indicator", xlSum
End Sub

```

----

The answer to your problem is [located here](http://www.mrexcel.com/forum/excel-questions/604666-visual-basic-applications-pivot-table-error-5-0_o.html#post2998219).

Your sheet name in `TableDestination:="LED OTTR!R1C1"` **needs to be surrounded with single quotes** in order for it to work `TableDestination:="'LED OTTR'!R1C1"`

You will also have problems with the duplicated name if you do not delete this PivotTable before rerunning the code.
