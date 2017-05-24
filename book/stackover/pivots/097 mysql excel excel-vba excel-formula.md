# SO item 097
I am trying to create a dynamic table - I have tried a Pivot Table, but cannot get it to work. So I thought that maybe it could be done with an IF-statement, but that did not work for me neither.

Basically, I have 2 tables, 1 table containing the information (data source table) and 1 table that should be dynamic according to the data in the first table.

![Table 1](https://i.stack.imgur.com/atIAv.gif)

So if I change the data in the E-column, the Fruit table (image below) must be updated accordingly.

![Table 2](https://i.stack.imgur.com/grVqB.gif)

So if I write 2 instead of 1 in the count of Apples, then it should create 2 apples under the "Fruit"-column". Data in the remaining columns will be calculated with a formula/fixed data - so that is not important.
I am open to any solutions; formulas, pivot tables, VBA, etc.

Have a nice weekend. I have both Excel 2010 and 2013.

----

If you want to repeat some text a number of times you can use a somewhat complicated formula to do it. It relies on there not being duplicate entries in the `Fruits` table and no entries with 0 count.

**Picture of ranges and results**

![ranges and results](https://i.stack.imgur.com/syPpQ.png)

**Formulas** involved include a starter cell `E2` and a repeating entry `E3` and copied down. These are actually _normal_ formulas, no array required. Note that I have created a `Table` for the data which allows me to use named fields to get the whole column.

```
E2 = INDEX(Table1[Fruits],1)
E3 = IF(
      INDEX(Table1[Count],MATCH(E2,Table1[Fruits],0))>COUNTIF($E$2:E2,E2),
      E2,
      INDEX(Table1[Fruits],MATCH(E2,Table1[Fruits],0)+1))

```

**How it works** This formula relies on checking the number of entries above the current one and comparing to the desired count. Some notes:

*   The starter cell is needed to get the first result.
*   After the first cell, it counts how often the value above appears in the total list. This is compared to the desired count. If less than desired, it will repeat the value from above. If greater, it will go to the next item in the list. There is a dual relative/absolute reference in here to count cells above.
*   Since it goes to the next item in the list, don't put a 0 for a count or it will get included once.

You can copy this down for as many cells as you want. It will `#REF!` when it runs out of data. You can wrap in an `IFERROR(..., "")` to make these display pretty.

If the non-0 rule is too much, it can probably be removed with a little effort. If there are duplicates, that will be much harder to deal with.
