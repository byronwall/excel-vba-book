# SO item 016
I need help with the following :

I have created a table on Sheet1 from A:E with column Titles from A1:E1 The titles are names , under each name there are numbers for example Carlos is the header title Located on A1 and under him there are numbers like 2 on cell A2 3 on cell A4 10 on cell A5

On Sheet2 I have the names for example from G4:G8 and I want to know the sum of the numbers depending on the name for example next to Carlos on H4 I should be getting 15 which is the sum of the values under his name on the Sheet1

I have tried vlookup - Index - HLOOKUP - conditional combining vlookp or Index but nothing works : (

----

I assume by "Table" you mean an actual named Table. In that case you can take advantage of this and use a formula like:

```
=SUM(INDEX(Table1, , MATCH(G4,Table1[#Headers],0)))

```

In my case, my table is named Table1 since it is a default name.

This formula works by searching through the `Table1[#Headers]` for the name in cell `G4`. It then use that column index to return an entire column using `INDEX`. Note that there is an empty `rownum` parameter to `INDEX` so it returns the whole column. From there it takes the `SUM` of this column.
