# SO item 017
**The Problem:**

I have a workbook with 3 sheets, each entitled "HeatNumbers", "HeatSheetTemplate", and "Heat vs Order". The Heat vs Order sheet has a number of new rows of data added to it daily, so the number of rows is always changing. Here is a picture of the column headings and some data:

![enter image description here](https://i.stack.imgur.com/nkyGJ.jpg)

**What I am looking for:**

On the HeatNumbers sheet, I have a button that executes some VBA code. Here is a pic of that sheet:

![enter image description here](https://i.stack.imgur.com/9ur7t.jpg)

Here is what I need to happen: A user will enter data into the black box in column J on a number of rows. Each line could contain an FO#. When the button is clicked, I need to filter all of the data on the Heat vs Order sheet above by any FO# in that black box region, copy that resultset over to the HeatNumbers sheet, beginning in row 2 col A, and then remove the filter from the Heat vs Order sheet.

**What I have tried:**

The only way I have been able to achieve this is by having the user manually filter the data on the Heat vs. Order sheet and copy and pasting the result to the HeatNumbers tab. This is cumbersome and prone to error, unfortunately.

Here is the code that was generated using the macro recorder:

```
Sub Filter_FO()
'
' Filter_FO Macro
'
    Range("A1:H20000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= Sheets("HeatNumbers").Range("J4:J22"), Unique:=False
    ActiveWindow.SmallScroll Down:=-15
    Range("A4:H300").Select
    Selection.Copy
    Sheets("HeatNumbers").Select
    ActiveWindow.SmallScroll Down:=-15
    Range("A2:H300").Select
    ActiveSheet.Paste
End Sub

```

----

In order to get the filter to work correctly, you need to use a `CriteriaRange` that only includes cells with values in them. The easiest way to do this is using the `.End(xlDown)` function. That works functions in the same way as CTRL+DOWN arrow and will select the last cell in that column of data so long as there are no blanks in between.

The second piece is to possibly improve the range of the data being filtered. Right now you are selecting a large area and hoping it includes the data you want. If your code works now, you can leave it. Improvements include:

*   Using just the column letters so that the whole columns are filtered. This works if the headers stay in row 1\. This is a little slow though.
*   If the data is a large block, you can use `.End(xlUp)` to find the last row and use that. This is included below.

And then the final piece is selecting the right range of data to copy. I just took the data range and selected visible cells using `.SpecialCells(xlCellTypeVisible)`.

In order for the copy to work cleanly, I clear out columns `A:H` on `HeatNumbers` to prevent any old data from sitting around. When I paste the data back over, I include the headers. This is the only real difference from your macro.

```
Sub FilterDataAndClearAndCopy()

    'get references to sheets
    Dim sht_data As Worksheet
    Dim sht_filter As Worksheet

    Set sht_data = Sheets("Heat vs Order")
    Set sht_filter = Sheets("HeatNumbers")

    'get the block of data to set the filter over
    Dim rng_data As Range
    Dim int_lastRow As Integer

    int_lastRow = sht_data.Range("A" & sht_data.Rows.Count).End(xlUp).Row
    Set rng_data = sht_data.Range("A1:H" & int_lastRow)

    'get the criteria range... assumes at least one entry below J3
    Dim rng_filter As Range
    Set rng_filter = Range(sht_filter.Range("J3"), sht_filter.Range("J3").End(xlDown))

    'filter the data
    rng_data.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=rng_filter, Unique:=False

    'clear out data
    sht_filter.Range("A:H").Clear

    'select data to copy
    rng_data.SpecialCells(xlCellTypeVisible).Copy

    'paste that data to filter sheet
    sht_filter.Range("A1").PasteSpecial xlPasteAll

    'remove the filter
    sht_data.ShowAllData

End Sub

```
