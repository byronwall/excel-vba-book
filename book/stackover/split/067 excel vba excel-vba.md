# SO item 067
I have an Excel spreadsheet of data for work that I need to split up in VBA. A couple of columns have multiple lines of text and others do not. I've figured out how to split the multiple lines of text, my problem is taking the column with a single line of text and copying it down. For example:

```
Company_Name     Drug_1      Phase_2        USA
                 Drug_2      Discontinued 
                 Drug_3      Phase_1        Europe
                 Drug_4      Discontinued  

```

Below is the code I am using to split columns B & C and then I can handle D manually, however I need column A to copy down into rows 2-4\. There's over 600 rows like this otherwise I would just do it manually. (Note: I'm putting column B into A below, and column C into C)

```
Sub Splitter()
    Dim iPtr1 As Integer
    Dim iPtr2 As Integer
    Dim iBreak As Integer
    Dim myVar As Integer
    Dim strTemp As String
    Dim iRow As Integer

'column A loop
    iRow = 0
    For iPtr1 = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        strTemp = Cells(iPtr1, 1)
        iBreak = InStr(strTemp, vbLf)
        Range("C1").Value = iBreak
            Do Until iBreak = 0
            If Len(Trim(Left(strTemp, iBreak - 1))) > 0 Then
                iRow = iRow + 1
                Cells(iRow, 2) = Left(strTemp, iBreak - 1)
            End If
            strTemp = Mid(strTemp, iBreak + 1)
            iBreak = InStr(strTemp, vbLf)
        Loop
        If Len(Trim(strTemp)) > 0 Then
            iRow = iRow + 1
            Cells(iRow, 2) = strTemp
        End If
    Next iPtr1

'column C loop
    iRow = 0
    For iPtr2 = 1 To Cells(Rows.Count, 3).End(xlUp).Row
        strTemp = Cells(iPtr2, 3)
        iBreak = InStr(strTemp, vbLf)
        Do Until iBreak = 0
            If Len(Trim(Left(strTemp, iBreak - 1))) > 0 Then
                iRow = iRow + 1
                Cells(iRow, 4) = Left(strTemp, iBreak - 1)
            End If
            strTemp = Mid(strTemp, iBreak + 1)
            iBreak = InStr(strTemp, vbLf)
        Loop
        If Len(Trim(strTemp)) > 0 Then
            iRow = iRow + 1
            Cells(iRow, 4) = strTemp
        End If
    Next iPtr2

End Sub

```

----

There is a bit of code I call the "waterfall fill" which does exactly this. If you can build a range of cells to fill (i.e. set `rng_in`), it will do it. It works on any number of columns which is a nice feature. You can honestly feed it a range of `A:D` and it will polish off your blanks.

```
Sub FillValueDown()

    Dim rng_in As Range
    Set rng_in = Range("B:B")

    On Error Resume Next

        Dim rng_cell As Range
        For Each rng_cell In rng_in.SpecialCells(xlCellTypeBlanks)
            rng_cell = rng_cell.End(xlUp)
        Next rng_cell

    On Error GoTo 0

End Sub

```

**Before and after**, shows the code filling down.

![enter image description here](https://i.stack.imgur.com/HEewM.png) ![enter image description here](https://i.stack.imgur.com/FzRce.png)

**How it works**

This code works by getting a range of all the blank cells. By default `SpecialCells` only looks into the `UsedRange` because of a [quirk with `xlCellTypeBlanks`](http://www.mrexcel.com/forum/excel-questions/371987-problem-specialcells-xlcelltypeblanks.html#post1844081). From there it sets the value of the blank cell equal to the closest cell on top of it using `End(xlUp)`. The error handling is in place because `xlCellTypeBlanks` will return an error if nothing is found. If you do the whole column with a blank row at top though (like the picture), the error will never get triggered.
