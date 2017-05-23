# SO item 053
I have the following data structure:

![Data+Expected Results+What I get](https://i.stack.imgur.com/xUq7X.png)

As you see in **column J**, I am trying to merge data into one column from columns **A** & **C** & **E** & **G**.

I am using this formula:

```
=IF(ROW()<=COUNTA($A:$A);INDEX($A:$C;ROW();COLUMN(A1));INDEX($A:$C;ROW()-COUNTA($A:$A)+1;COLUMN(C1)))

```

and I get the values in column **K** as you see. Currently this formula is merging only two columns. **How to modify it to merge all four columns?**

And how to only get those values starting from **row 5**?
The column height will vary constantly: sometimes there are 10 values in column A and sometimes there are 2 values.

_Either any excel formula or any VBA code will be acceptable._

----

This answer is another way of thinking about the formulas you could use for this sort of task. It gets to the point made by @Jeeped that it is difficult to find unique values in multiple columns. My first step then is to create a single column.

If you can live with a helper column, these formulas might be a tad easier to maintain than the nested `IFERROR` already proposed. They are equally difficult to understand though at first glance. The other upside is that it scales nicely if the number of columns involved increases.

It is possible using `CHOOSE` and some `INDEX` math to build a single column array of a group of separated columns. The trick is that `CHOOSE` will join discontinuous ranges side-by-side when given an array as the selecting parameter. If this starts with columns of the same size, you can then use division and mod math to turn it into a single column.

**Picture of ranges** shows the four groups of data with duplicates colored red.

![picture of ranges](https://i.stack.imgur.com/jaPIP.png)

**Formula** in `F2:F31` is an array formula. This is combining all of the columns into an array and then back into a single column. I selected the columns out of order just to emphasize that it is handling a discontinuous range.

```
=INDEX(CHOOSE({1,2,3,4}, A2:A7,C2:C7,B2:B7,D2:D7), MOD(ROW(1:30)-1, ROWS(A2:A7))+1,INT((ROW(1:30)-1)/ROWS(A2:A7))+1)

```

The array formula in `H2` and copied down is then the standard formula for unique values. The one exception is that instead of avoiding blanks like normal, I am avoiding 0 values.

```
=IFERROR(INDEX(F2:F31,MATCH(0,IF(F2:F31=0,1,COUNTIF($H$1:H1,F2:F31)),0)),"")

```

A couple of other comments about this approach:

*   In the `CHOOSE`, I am using `{1,2,3,4}`. This could be replaced with `TRANSPOSE(ROWS(1:4))` or whatever number of columns you have.
*   There is also a `ROWS(A2:A7)` in 2 places, this could just be `2:7` or `1:6` or whatever size was used for the column size. I used one of the data ranges so that the coloring was simplified and to emphasize it needs to match the size of the block.
*   And the `ROW(1:30)` is used for the number of total items to collect. It really only needs to be `1:24` since there are `6*4` items, but I made it big while testing.

There are definitely a couple of downsides to this approach, but it may be a good trick to keep in the toolbox. Never know when you might want to make a column out of discontinuous ranges. The largest downside is that the columns of data all need to be the same size (and of course the helper column).
