# SO item 105
Here is what I am trying to do, I have a sheet that has a list of names with no duplicates that varies in length. I would like to have either with a formula or vba sub, to have the next row copy the original list of names and offset it my one, so that the top name on the original is now the last name of the second list. I need to have at the end 10 list where none of the same names are in the same row.

Here is a sample of what I'd like it to look like.

```
ColumnB   ColumnC   ColumnD   ColumnE
Name1     Name2     Name3     Name4
Name2     Name3     Name4     Name1
Name3     Name4     Name1     Name2
Name4     Name1     Name2     Name3

```

Like a game of Sudoku, none of the names in each row or column can have a duplicate.

I am not sure how to best achieve this since as mentioned above the length of the list is a variable. Ideally I'd like to create the first list, then have the other 9 list to auto populate. Any suggestions?

EDIT___________________ @Paul Drye, I get the following results with your formula

```
ColumnB   ColumnC   ColumnD   ColumnE
Name1     Name2     Name3     Name4
Name2     Name3     Name4     Name1
Name3     Name4     Name1     Name1
Name4     Name1     Name1     Name1

```

As you can see, the last two columns start showing an issue.

----

If a formula works, you can get this result simply with

```
=IF(ISBLANK(B3),B$2,B3)

```

in cell `C2`, assuming your data starts in `B2`. This can then be copied down and over or filled using <kbd>CTRL+R, CTRL+D</kbd> after selecting the whole range of cells to occupy.

If the copy is correct, the formula of cell `K11` is:

```
=IF(ISBLANK(J12),J$2,J12)

```

**Picture** shows the inputs in column B, the rest are this formula

![results of formula](https://i.stack.imgur.com/xCZ3m.png)

This formula works more or less because of the absolute row reference which ensures that the value from row 2 is used if we are at the end of the list.
