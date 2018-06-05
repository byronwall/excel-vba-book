# SO item 057
I have a table that is pulling thousands of rows of data from a very large sheet. Some of the columns in the table are getting their data from every 5th row on that large sheet. In order to speed up the process of creating the cell references, I used an OFFSET formula to grab a cell from every 5th row:

```
=OFFSET('Large Sheet'!B$2572,(ROW(1:1)-1)*5,,)
=OFFSET('Large Sheet'!B$2572,(ROW(2:2)-1)*5,,)
=OFFSET('Large Sheet'!B$2572,(ROW(3:3)-1)*5,,)
=OFFSET('Large Sheet'!B$2572,(ROW(4:4)-1)*5,,)
=OFFSET('Large Sheet'!B$2572,(ROW(5:5)-1)*5,,)
etc...

```

OFFSET can eat up resources during calculation of large tables though, and I'm looking for a way to speed up/simplify my formula. Is there any easy way to convert the OFFSET formula into just a simple cell reference like:

```
='Large Sheet'!B2572
='Large Sheet'!B2577
='Large Sheet'!B2582
='Large Sheet'!B2587
='Large Sheet'!B2592
etc...

```

I can't just paste values either. This needs to be an active reference, because the large sheet will change.

Thanks for your help.

----

And here is one last approach to this that does not use VBA or formulas. It's just a quick and dirty use of AutoFilter and deleting rows.

**Main idea**

*   Add a reference to a cell `=Sheet1!A1` and copy it down to match as many rows as there are in the main data.
*   Add another formula in `B1` to be `=MOD(ROW(), 5)`
*   Filter column `B` and uncheck the 0s (or any single number)
*   Delete all the rows that are visible
*   Delete column B
*   Voila, formulas for every 5th row

**Some reference images**, these are all taken on `Sheet2`.

Formulas with AutoFilter ready.

![formulas with filter](https://i.stack.imgur.com/KNC8G.png)

Filtered and ready to delete

![filtered](https://i.stack.imgur.com/Ihv8p.png)

Delete all those rows (select `A1`, CTRL+SHIFT+DOWN ARROW, SHIFT+SPACE, CTRL+MINUS)

![delete rows](https://i.stack.imgur.com/4jjCZ.png)

Delete column B to get final result with "pure" formulas every 5th row.

![result](https://i.stack.imgur.com/DqOnV.png)
