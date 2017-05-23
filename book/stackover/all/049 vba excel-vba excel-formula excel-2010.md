# SO item 049
Google Doc with data in its current format followed by desired format: [https://docs.google.com/spreadsheets/d/1XlxEVcP6QpWYyOeflLmp_mKflCBclim_UQSeMkHByh8/edit?usp=sharing](https://docs.google.com/spreadsheets/d/1XlxEVcP6QpWYyOeflLmp_mKflCBclim_UQSeMkHByh8/edit?usp=sharing)

I am trying to create a template to rearrange a data set that is exported in a horrible format. I have posted a link to a Google Doc that has an example of the data in its current format followed by how I need it to be formatted. Currently, all data for a given person is in a single row, and the ID # is repeated before each record, as shown. Each record consists of 12 columns, and this is repeated 31 times across a row, totaling 372 cells per row. There are 838 rows (or 837 without the headers). I need either a set of formulas or a macro that will separate a single row of data (the 372 cells) into 31 rows of 12 columns for an entire spreadsheet. I have been able to accomplish this only with a single row (using the `offset` function and then again using `index`), but I am struggling to find a way to make it apply to every row on a worksheet. Once that first row is done, the formula goes no further. Ideally, the rearranged data will appear on a **new** worksheet. I can't just manually separate the rows and filter them by ID #, because then I would have to redo that every time I rerun the report. Please let me know if I can give any further clarification!

----

Here is a formula based solution. It assumes that you can create a new sheet to reform the data. Based on your description, sounds like this is what you want anyways. I have dummied down the example to have 4 categories and 3 repeats per row. Change the 4 and 3 below to match your 12 and 31\. (Harder to take a snapshot like that!)

**Picture of data and results.**

Data. You can pretend that my column header "A" is your "ID".

![data](https://i.stack.imgur.com/3MbdC.png)

Results. It repeats the headers for simplicity. You can delete those out at the end.

![results](https://i.stack.imgur.com/RzThk.png)

**Formula** in `A1` on `Sheet2` and copied over 4 columns and down as many rows as needed.

```
=INDEX(Sheet1!$A$1:$L$7,INT((ROW()-1)/3)+1,MOD(ROW()-1,3)*4+COLUMN())

```

**How it works**

*   `INDEX` is used to return a given cell from the array of the entire data
*   `INDEX` needs a row and column to retrieve so I used integer division to determine the row. We know that a given row in the results needs to be repeated for as many repeats of the headers as there are. In this case 3\. The `ROW()` refers to the row in the results, and the `-1` is to ensure that it starts at `0` instead of `1`.
*   The column for `INDEX` comes from the same idea. For a given row, it needs to get a column that is the current column "pushed" over by the repeating row. The `MOD` here ensures that the column numbers repeat in a small range even though the row is becoming a large number.

Finally, this formula really relies on the results starting in `A1` on a fresh sheet. You can do it differently, but the formulas will become (even more of) a mess.
