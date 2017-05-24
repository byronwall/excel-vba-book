# SO item 055
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

Here is one approach using `INDEX` instead of `OFFSET`. I am not sure if it is faster, I guess you can check. `INDEX` is not volatile, so you might get some advantage from that.

**Picture of ranges**, you can see that `Sheet1` has a lot of data and `Sheet2` is pulling every 5th row from that sheet. The data in `Sheet1` goes from `A1:F1000` and just reports the address of the current cell.

![sheets](https://i.stack.imgur.com/tQlSp.png)

**Formulas** use `INDEX` and are copied down and across from `A1` on `Sheet2`.

```
=INDEX(Sheet1!$A$1:$F$1000,ROW()*5,COLUMN())

```
