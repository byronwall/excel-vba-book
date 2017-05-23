# SO item 103
I have 2 different files which have different headers, for example:

```
OldfileHeaders | NewFileheaders
ID             | Test ID
Date           | New date

```

and so on. I am trying to compare the data in both sheets and see if they match. The rows of data may be in different order and the headers may also be in different order.

So what I am trying to do is: 1) define which headers match which headers between the 2 files 2) find the ID from the oldfile and see if it is in the new file, if it is then see if the data under each header matches. If it doesn't then export that row of data to a new sheet add a column and label it "Missing".

The Code So far:

```
Set testIdData = testIdData.Resize(testIdData.CurrentRegion.Rows.Count)

Do Until sourceId.Value = ""
    datacopy = False
    ' Look for ID in test data
    Set cellFound = testIdData.Find(What:=sourceId.Value, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If cellFound Is Nothing Then
    ' This entry not found, so copy to output
        datacopy = True
        outputRange.Resize(ColumnSize:=NUMCOLUMNS).Interior.Color = vbRed
    Else
        ' This assumes that columns are in same order
        For columnNum = 2 To NUM_COLUMNS_DATA
        ' No need to test the ID column
            If sourceId.Cells(ColumnIndex:=columnNum).Value <> cellFound.Cells(ColumnIndex:=columnNum).Value Then
                outputRange.Cells(ColumnIndex:=columnNum).Interior.Color = vbYellow
                datacopy = True
            End If
        Next columnNum
    End If
    If datacopy Then
        sourceId.Resize(ColumnSize:=NUMCOLUMNS).Copy
        outputRange.PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        Set outputRange = outputRange.Offset(RowOffset:=1)
        difference = difference + 1
    End If
    Set sourceId = sourceId.Offset(RowOffset:=1)
Loop

```

This code works depending on me formatting the sheets in the correct order and changing the header names.

I need help in defining which field names match which field names within the 2 sheets, and then searching the new sheet for each ID and seeing if the data in the corresponding cells match. If the ID is not in the sheet then output that row too a different sheet. If the id is present and there are differences in the cells then out put these to the shame sheet. I want to produce a tally of differences in each column.

----

Matching up data between data sets requires that you give the program some help. In this case, the help needed is which columns are related to each other. You have identified a small table of how headers are related. With this, you can do the various translations from data source 1 to data source 2\. It requires heavy usage of `Application.Match` and `Application.VLookup`.

I will provide a base example which does the core of what you are trying to do. It is much easier to see it all on one sheet which is what I have done.

**Picture of data** shows three tables: rng_headers, rng_source, and rng_dest. One is the lookup for the headers, the second is the "source" data, and the third is the data source to compare against which I will call destination = "dest".

![starting data](https://i.stack.imgur.com/Z9xld.png)

**Code** include steps to: iterate through all the IDs in the source data, check if they exist in the dest data, and, if so, check all the individual values for equality. This code checks the headers on every step (which is slow) but allows for the data to be out of order.

```
Sub ConfirmHeadersAndMatch()

    Dim rng_headers As Range
    Set rng_headers = Range("B3").CurrentRegion

    Dim rng_dest As Range
    Set rng_dest = Range("I2").CurrentRegion

    Dim rng_source As Range
    Set rng_source = Range("E2").CurrentRegion

    Dim rng_id As Range 'first column, below header row
    For Each rng_id In Intersect(rng_source.Columns(1).Offset(1), rng_source)

        Dim str_header As Variant
        str_header = Application.VLookup( _
            Intersect(rng_id.EntireColumn, rng_source.Rows(1)), _
            rng_headers, 2, False)

        'get col number
        Dim int_col_id As Integer
        int_col_id = Application.Match(str_header, rng_dest.Rows(1), 0)

        'find ID in the new column
        Dim int_row_id As Variant
        int_row_id = Application.Match(rng_id, rng_dest.Columns(int_col_id), 0)

        If IsError(int_row_id) Then
            'ID missing... do something
            rng_id.Interior.Color = 255
        Else
            Dim rng_check As Range 'all values, same row
            For Each rng_check In Intersect(rng_source, rng_id.EntireRow)

                'get col number
                str_header = Application.VLookup( _
                    Intersect(rng_check.EntireColumn, rng_source.Rows(1)), _
                    rng_headers, 2, False)
                int_col_id = Application.Match(str_header, rng_dest.Rows(1), 0)

                'check value
                If rng_check.Value <> rng_dest.Cells(int_row_id, int_col_id).Value Then
                    'values did not match... do something
                    rng_dest.Cells(int_row_id, int_col_id).Interior.Color = 255
                End If

            Next rng_check
        End If
    Next
End Sub

```

**Notes on the code**

*   Ranges are built on `CurrentRegion` which picks out the blocks of data. You can swap these out for different ranges on different sheets.
*   Column header translation is done with `Application.VLookup` to check the source header and return the destination header. This `String` is then found in the destination header row using `Application.Match`. You could abstract this code into a `Function` to avoid repeating it twice.
*   Once the column is found, the ID is searched for in the destination table using `Application.Match`. This will return an error if the ID is not found.
*   If the ID is found, it then checks all of the other values in the same row, comparing them against the correct columns in the destination table. Non-matching results are colored red.
*   If all of the columns do not have pairs, you can add additional checks on the `VLookup` or the column `Match` to check this.
*   The vast majority of this code just handles getting to the correct spots in the data using `Intersect`, `Rows`, and `Columns`.

**Results** show some red values for the ID not found and the values that don't match.

![results](https://i.stack.imgur.com/9xSWf.png)
