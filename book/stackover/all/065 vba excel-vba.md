# SO item 065
Been working on this AutoFiltering code for a while. It works well as far as it goes. If I use my search criteria in "Quotes" replacing FilterCriteria it works everytime. However, when trying to pass numbers along in FilterCriteria it fails to find anything in my range (A:D ONLY!) everytime. It finds all text fields in Colums E:G fine as they are all text. Columns A:D returns nothing. I tried formatting A:D as text instead of numbers and it STILL sees nothing when filtering. Sample range shown hopefully at end.

```
Sub FindProduct()

  'Note: This macro uses the function LastRow at end of Module
  ' Highly moded code from Ron de Bruin

    'To define My_Range
       Dim My_Range As Range
       Dim CalcMode As Long
       Dim ViewMode As Long
       Dim CCount As Long
    'To define New Sheet and Range
       Dim WSNew As Worksheet
    'Use for column and filter data selection
       Dim FilterCriteria As String
       Dim PickCol As String

    'Set filter range on ActiveSheet
       Set My_Range = Range("A1:G" & LastRow(ActiveSheet))
       My_Range.Parent.Select

 '  ************************************
    My_Range.Parent.AutoFilterMode = False
       '  Unprotect sheet, turn off AutoFilter, Show All
          With ActiveSheet
             .Unprotect
             On Error Resume Next
             .ShowAllData
          End With
    '  Code to check if workbook is protected here. Redundant.
 '  ****************************************
     'Turn off ScreenUpdating, Calculation, EnableEvents code here
  '  +++++++++++++++++++++++++++++++++++
       '  Use this to pick a Column to search and your FilterCriteria
       PickCol = InputBox("What Column do you want to search in " & vbCrLf _
       & "(A=1,B=2,C=3,D=4,E=5,F=6,G=7)?" _
       & vbCrLf & vbCrLf, "Select Column to Search")
          '  Input error check
       '  ######################
       FilterCriteria = InputBox("What are you looking for?" _
       & vbCrLf & vbCrLf & "This will work with partial Information.", _
       "Enter Filter Parameter")
          '  Input error check
 '  *********************************************************
    '  Insert PickCol and FilterCriteria variables
    My_Range.AutoFilter Field:=PickCol, Criteria1:="=*" & FilterCriteria & "*"

    'Check if there are not more then 8192 areas (limit of areas that Excel can copy)
    CCount = 0
    On Error Resume Next
    CCount = My_Range.Columns(1).SpecialCells(xlCellTypeVisible).Areas(1).Cells.Count
    On Error GoTo 0
      If CCount = 0 Then
          MsgBox "There are more than 8192 areas:" _
               & vbCrLf & "It is not possible to copy the visible data."
      Else
        '  ***********************************************
           'Delete "Filtered Data" sheet if it exists code here
        '  ***********************************************
        '  ------------------------------
          'Add a new Worksheet
           Set WSNew = Worksheets.Add(After:=Sheets(ActiveSheet.Index))
           On Error Resume Next
           WSNew.Name = "Filtered Data"
        '  ------------------------------
        '  ///////////////////////////////////////////////////
           'Copy/paste the visible data to the new worksheet
           My_Range.Parent.AutoFilter.Range.Copy
             ' Paste copied range starting at Cell("A2")
             With WSNew.Range("A2")
                 .PasteSpecial Paste:=8
                 .PasteSpecial xlPasteAll
                 .PasteSpecial xlPasteFormats
                 Application.CutCopyMode = False
                 .Select
             End With
        ' ///////////////////////////////////////////////////
        ' *****************************************
          'Adds Formatted Text to Cell ("A1") code here
        ' *****************************************
      End If

    ' Turn off AutoFilter
    My_Range.Parent.AutoFilterMode = False

'  ******************************************************
   'More finishing code here
'  ******************************************************

End Sub

 Function LastRow(Sh As Worksheet)
     On Error Resume Next
     LastRow = Sh.Cells.Find(What:="*", _
                        After:=Sh.Range("A1"), _
                        Lookat:=xlPart, _
                        LookIn:=xlValues, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
     On Error GoTo 0
 End Function

```

Sample data:

```
    A        B     C         D          E         F           G
Date Rvd    Qty   File#    P.O.#    Cust Name    Vend Name   Carrier
02/14/15    210   41680    38565    Some Tech    John        DHL
03/08/15    458   17017    38569    Them Guys    Donn        Fedx
03/12/15    350   16736    38541    Some Guys    Teri        UPS
03/24/15    236   42630    38655    Some Tech    John        DHL
04/08/15    458   56985    85693    Them Guys    Donn        Fedx
04/12/15    350   12345    43851    Some Guys    Teri        UPS
04/18/15    838   56685    85693    Them Guys    Donn        Fedx
05/05/15    110   13245    43851    Some Guys    Teri        UPS

```

For whatever reason when it runs the AutoFilter using any numbers for A:D it fails to give any filtered Data. I'm stumped as I said It WILL return filtered data IF I place the exact value I want in the AutoFilter line.

Pretty sure this line is my issue/ problem: **My_Range.AutoFilter Field:=PickCol, Criteria1:="=_" & FilterCriteria & "_"**

Any Ideas?

I guess now I have to figure out how to actually make that work. Using Autofilter properly on the sheet it works fine. If I have to do as I think the article shows then I have to add 4 more columns AND I have to rewrite the code in the SaveLog Code on the form that generates this list. Sounds like I need to substantially increase the size of my code for everything. For a Novice as myself I'm certainly overwhelmed at this point.

----

The core of this issue is that you cannot use Text comparison operators with Numbers. **When you add the wildcards `*` to the search criteria, you are enforcing a Text comparison.**

If you want this to work with numbers and text and have the variable column selection, you will need to add some checks to build the criteria correctly. This would involve dropping the `*` when a number column is selected. The main thing to keep in mind is that each data type only has certain filters that are available to it. To check those, click the arrow in the normal filter menu to see what is listed under `Number Filters` or `Date Filters` or `Text Filters`.

Given all of that, **if you want to filter those numerical columns on `Contains`, you will need to convert it to text.**

Per the [comment by @Tim Williams](http://stackoverflow.com/questions/30520060/autofilter-not-seeing-numerical-data-in-filtered-view-multiple-values#comment49119082_30520060), you can convert your numbers to text using the `Data->Text to Columns` feature. You can automate this step using VBA if you know which ranges need to be converted.

The minimum number of parameters required to get that to work appears to be `DataType` and `FieldInfo`. `FieldInfo` is the important one for forcing the conversion.

```
Sub ConvertColumnNumberToText()

    Dim rng_column As Range
    For Each rng_column In Range("B:D").Columns
        rng_column.TextToColumns DataType:=xlDelimited, FieldInfo:=Array(1, 2)
    Next rng_column

End Sub

```

Check the documentation for [TextToColumns](https://msdn.microsoft.com/en-us/library/office/ff193593.aspx) to see what the parameters are. It will only work on a column at a time, hence the loop.

Also, there is little harm in running this code multiple times, so long as it only runs on columns with numbers only. If you accidentally run it on a column that can be split into columns (contains a `TAB` by default), you will start to overwrite other columns.
