# SO item 052
I'm having trouble with my hyperlinks.

I have some code which works very well for its purpose. However, it does something additionally which I don't want and to me makes little sense why it is happening.

The code I have is below:

```
Dim hLink As Hyperlink
Dim cColumn As Range
Dim Path1 As String
Dim Path2 As String
Dim pathEnd As Integer

Set cColumn = Columns(int4)
str3 = ColumnLetter(ActiveCell.Column)

For Each hLink In cColumn.Hyperlinks
    pathEnd = InStr(hLink.SubAddress, "!")
    Path1 = Left(hLink.SubAddress, pathEnd)
    pathEnd = Len(hLink.SubAddress) - InStr(hLink.SubAddress, ColLetter)
    Path2 = Right(hLink.SubAddress, pathEnd)
    hLink.SubAddress = Path1 & str3 & Path2
Next hLink

```

int1 finds the column number in a previous subroutine. ColumnLetter finds the column letter of the new column.

Here's what the full code does (some of which isn't included here).

I have a "template" column which is copied to a new column. The info is updated in the new column and then towards the end of the programme, this hyperlink subroutine is run.

It works very well, but replacing the letter of the template column within the hyperlink address to the new column.

However once it has run, the hyperlinks in the template column have also changed.

I have stopped the code before the hyperlink subroutine is run and the hyperlinks are as expected and have not been changed - ie both columns match the links from the template column. Therefore I am confident this is the problem code (which makes sense).

I have tried a number of iterations of selecting the new column, to no avail, it always changes the hyperlinks in both columns.

I have even manually run through the code using F8, checking the column number and row number of each hyperlink it seems to be updating and it doesn't even change to the template column!

I'm at a loss. Please help.

----

I believe `Hyperlinks` are stored at a level above the `Range` but can be returned at the `Range` level. That is, Excel is storing all of the `Hyperlinks` in one place and then gives you a convenient function to return the `Hyperlinks` that are anchored in a given `Range`. You will find some real oddities if you output all of the info for the Hyperlinks after copy/pasting them around. It looks like if you copy a range, you just copy a reference to the Hyperlink and do not get a new one.

I think if you want to change the Hyperlink and not affect the other one, **you probably need to create a new Hyperlink.**

**Code to show some of the oddities**

```
Sub MakeHyperlinks()

    Dim rng_cell1 As Range
    Set rng_cell1 = Range("A1")

    'create a hyperlink to cell one row below
    rng_cell1.Hyperlinks.Add rng_cell1, "", rng_cell1.Offset(1).Address, , rng_cell1.Offset(1).Address

    'copy that column and paste (insert) next door several times
    For i = 1 To 5
        rng_cell1.EntireColumn.Copy
        rng_cell1.EntireColumn.Insert
    Next

    OutputHyperlinkInfo

    'change original hyperlink
    rng_cell1.Hyperlinks(1).SubAddress = "b2"

    OutputHyperlinkInfo

End Sub

Sub OutputHyperlinkInfo()

    Dim sht As Worksheet
    Set sht = ActiveSheet

    Dim hyp As Hyperlink
    Dim rng_hyp As Range

    Debug.Print Join(Array("rng.address", "hyp.Name", "hyp.Range", "hyp.Address", "hyp.SubAddress", "hyp.TextToDisplay"), "|")

    For Each rng_hyp In sht.UsedRange.SpecialCells(xlCellTypeConstants)
        For Each hyp In rng_hyp.Hyperlinks
            Debug.Print Join(Array(rng_hyp.Address, hyp.Name, hyp.Range, hyp.Address, hyp.SubAddress, hyp.TextToDisplay), "|")
        Next
    Next
End Sub

```

**Result** includes the Immediate output of the Subs.

```
rng.address|hyp.Name|hyp.Range|hyp.Address|hyp.SubAddress|hyp.TextToDisplay
$A$1|$A$2|$A$2||$A$2|$A$2
$B$1|$A$2|$A$2||$A$2|$A$2
$C$1|$A$2|$A$2||$A$2|$A$2
$D$1|$A$2|$A$2||$A$2|$A$2
$E$1|$A$2|$A$2||$A$2|$A$2
$F$1|$A$2|$A$2||$A$2|$A$2
rng.address|hyp.Name|hyp.Range|hyp.Address|hyp.SubAddress|hyp.TextToDisplay
$A$1|$A$2|$A$2||b2|$A$2
$B$1|$A$2|$A$2||b2|$A$2
$C$1|$A$2|$A$2||b2|$A$2
$D$1|$A$2|$A$2||b2|$A$2
$E$1|$A$2|$A$2||b2|$A$2
$F$1|$A$2|$A$2||b2|$A$2

```

The important thing to note here are that all of the `SubAddress` were changed even though the original call was to a single cell. It it also somewhat curious that all of the Hyperlinks have the same `Name`. Not sure if that is indicative of what is happening here.
