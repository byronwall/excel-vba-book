# SO item 121
This is the code which I used to copy data from multiple sheets to single sheet. I want to know if there is any way by which I can copy the data into "Report" sheet starting from 3rd Column, i.e, the data should be pasted into sheet from 3rd column onwards.

```
Sub AppendDataAfterLastColumn()
Dim sh As Worksheet
Dim DestSh As Worksheet
Dim Last As Variant
Dim CopyRng As Range

With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With

' Delete the summary worksheet if it exists.
Application.DisplayAlerts = False
On Error Resume Next
ActiveWorkbook.Worksheets("Report").Delete
On Error GoTo 0
Application.DisplayAlerts = True

' Add a worksheet with the name "Report"
Set DestSh = ActiveWorkbook.Worksheets.Add
DestSh.Name = "Report"

' Loop through all worksheets and copy the data to the
' summary worksheet.
For Each sh In ActiveWorkbook.Worksheets
    If sh.Name <> DestSh.Name Then
  lastcol = DestSh.Cells(1, DestSh.Columns.Count).End(xlToLeft).Column
        ' Find the last column with data on the summary
        ' worksheet.
        Last = lastcol
  lastCol3 = sh.Cells(1, sh.Columns.Count).End(xlToLeft).Column
        ' Fill in the columns that you want to copy.
        Set CopyRng = sh.Range(sh.Cells(1, 2), sh.Cells(15, lastCol3))

        ' Test to see whether there enough rows in the summary
        ' worksheet to copy all the data.
        If Last + CopyRng.Columns.Count > DestSh.Columns.Count Then
            MsgBox "There are not enough columns in " & _
               "the summary worksheet."
            GoTo ExitTheSub
        End If

        ' This statement copies values, formats, and the column width.
        CopyRng.Copy
        With DestSh.Cells(1, Last + 1)
            .PasteSpecial 8    ' Column width
            .PasteSpecial xlPasteValues

           '.PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
        End With

    End If
Next

ExitTheSub:

Application.Goto DestSh.Cells(1)

With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With
End Sub

```

Data sheet 1 from comments:

![enter image description here](https://i.stack.imgur.com/Szskq.png)

Data sheet 2 from comments:

![enter image description here](https://i.stack.imgur.com/RYlaf.png)

Expected output from comments:

![enter image description here](https://i.stack.imgur.com/3zp3c.png)

----

This sort of copy can be done easily with `Copy`. In order to pick the output `Range` for the paste part, you can use an `Application.InputBox` with a `Type:=8` parameter. This prompts Excel to open the `Range` selection dialog which works well.

Once you know those two pieces, the only difficulty is building the `Ranges`. This is not difficult, but _is_ specific to the context, existing data on the sheets, and degree of robustness. For the example below, I am using `CurrentRegion` to get the block of data (same as hitting <kbd>CTRL+A</kbd>) and then `Intersect` to only get the desired columns. You can also make use of `UsedRange` and `End` to build ranges.

**Picture of ranges** shows the different sheets for input and the final sheet for output. The sheet to paste into `c` is empty for now.

![data and empty sheet](https://i.stack.imgur.com/o5MbZ.png)

**Code** does the work to get the two ranges to copy and then prompts for an output location. From there, it pastes the resulting `Ranges` into the desired location. There is an `Offset` to ensure that the 2nd range does not overlap the first.

```
Sub CopyFromTwoRanges()

    Dim rng_set1 As Range
    Dim rng_set2 As Range

    Dim rng_output As Range

    'build the ranges
    Set rng_set1 = Intersect(Sheets("a").Range("C:F"), _
        Sheets("a").Range("C1").CurrentRegion)

    Set rng_set2 = Intersect(Sheets("b").Range("C:F"), _
        Sheets("b").Range("C1").CurrentRegion)

    'prompt for cell
    Set rng_output = Application.InputBox("Pick the range", Type:=8)

    'ensure a single cell only
    Set rng_output = rng_output.Cells(1, 1)

    'paste the ranges
    rng_set1.Copy rng_output
    rng_set2.Copy rng_output.Offset(, rng_set1.Columns.Count)

End Sub

```

**Result** shows the prompt with cell selected and then the output.

![enter image description here](https://i.stack.imgur.com/3JFJx.png)

![enter image description here](https://i.stack.imgur.com/UKkEh.png)
