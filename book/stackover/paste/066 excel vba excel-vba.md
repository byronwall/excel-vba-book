# SO item 066
See picture: [http://s12.postimg.org/ov8djtuh9/Capture.jpg](http://s12.postimg.org/ov8djtuh9/Capture.jpg)

**Context:** Trying to activate a sheet (variable: cSheet) in another workbook and paste data there from copied data from a different workbook. I'm getting a subscript out of range error whenever I try to activate directly using the variable (i.e. Worksheets(Name).Activate) or try to define a worksheet using the variable and then activate it. I've also tried other coding styles, using "With Worksheet" etc. and my code was a lot longer but I started over because every time I fix something, something else goes wrong. So, sticking to the basics. Any help would be greatly appreciated.

```
Sub GenSumRep()

Dim AutoSR As Workbook
Dim asrSheet As Worksheet
Dim tempWB As Workbook
Dim dataWB As Workbook
Dim SecName As String
Dim oldcell As String
Dim nsName As String
Dim cSheet As Worksheet

Set AutoSR = ActiveWorkbook
Set asrSheet = AutoSR.ActiveSheet

For a = 3 To 10

    SecName = asrSheet.Range("D" & a).Value

    If SecName <> "" Then

    Workbooks.Open Range("B" & a).Value
    Set tempWB = ActiveWorkbook
    'tempWB.Windows(1).Visible = False

    AutoSR.Activate

    Workbooks.Open Range("C" & a).Value
    Set dataWB = ActiveWorkbook
    'dataWB.Windows(1).Visible = False

    AutoSR.Activate

        'Copy paste data
        For b = 24 To 29
        oldcell = Range("C" & b).Value
            If b = 24 Then
            nsName = Trim(SecName) & " Data"
            Set cSheet = tempWB.Sheets(nsName)
            Else
            nsName = asrSheet.Range("B" & b).Value
            Set cSheet = tempWB.Sheets(nsName)
            End If

        'Copy
        dataWB.Activate
        Range(oldcell).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy

        'Paste
        tempWB.Activate
        cSheet.Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False

        b = b + 1
        Next b

    End If

a = a + 1

Next a

End Sub

```

----

You only get that error for one reason: **the name your provided does not exist in the collection!**

There are a couple of likely reasons for this based on your code:

*   Your `nsName` variable contains hidden characters that make it different even though it appears correct.
*   You are looking for the sheet in the wrong workbook.

Based on your comments, **it seems that you are looking in the wrong workbook**. A good way to check out these subscript errors is to iterate the collection and print out the `Names` included therein.

```
Dim sht as Worksheet    
For Each sht In tempWB.Sheets
    Debug.Print sht.Name
Next sht

```

In general, it is desirable to get rid of calls to `Select` and `Activate` so that you are not relying on the interface in order to get objects. See [this post about avoiding `Select` and `Activate`](http://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba-macros) for more info.

One idea applied to your code is to assign the Workbooks directly without `ActiveWorkbook`:

```
Set tempWB = Workbooks.Open(asrSheet.Range("B" & a).Value)
Set dataWB = Workbooks.Open(asrSheet.Range("C" & a).Value)

```

instead of:

```
    Workbooks.Open Range("B" & a).Value
    Set tempWB = ActiveWorkbook
    'tempWB.Windows(1).Visible = False

    AutoSR.Activate

    Workbooks.Open Range("C" & a).Value
    Set dataWB = ActiveWorkbook
    'dataWB.Windows(1).Visible = False

```
