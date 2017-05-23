# SO item 106
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

If you want a formula that works regardless of what surrounds your data, you can get the same result using `ROWS`, `COLUMNS`, and `MOD` along with some absolute/relative ranges.

Formula in cell `C2` copied down and over. Probably looks a little better with a named range. If you want to see how it generates the numbers, remove `INDEX` and get the counter.

```
=INDEX($B$2:$B$11,MOD(ROWS($C$2:C2)+COLUMNS($C$2:C2)-1,COUNTA($B$2:$B$11))+1)

```

**Picture** shows the same result as the other answer

![results and formula](https://i.stack.imgur.com/J0JPN.png)
