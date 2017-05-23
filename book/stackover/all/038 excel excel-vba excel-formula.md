# SO item 038
Please give a hand with this im trying to have a conditional formatting to apply to input cells in a spreadsheet but at the same time i want a switch to turn it off for printing purposes.

Formula:

```
=IF(READY_TO_PRINT="YES",CELL("protect",A1)=0)

```

The ready to print is a name range to serve as a switch to turn the specify style i want to turn off before print.

I would appreciate any solution including vba scenarios. Thanks in advance!

----

If you want to apply the formatting with both conditions, you just need an `AND` around them. Conditional formatting expects to get a `TRUE/FALSE` answer to apply the formatting. `AND` does this properly while `IF` will not as you've written it.

```
=AND(READY_TO_PRINT="NO",CELL("protect",A1)=0)

```

I switched your `READY_TO_PRINT` to `"NO"` since it seems you want to apply the formatting when it is not ready to print (as I understand it). If not, hopefully you can modify this formula as needed to get your solution.
