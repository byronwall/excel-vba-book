# SO item 081
I'm using Excel 2013 VBA. I have the following data

```
       A                    B
1   John  Doe       John  Doe
2   Mary  Smith     Mary  Smith
3   Alice  Jones    Alice  Jones
4   Bob C Carter    Bob  Carter
5   David L Macy    David L Macy
6   June  Weaver    June  Weaver

```

I'm searching for exact matches of the entries in Column B with the range of entries in Column A using the partial code:

```
Dim lastAllScriptsRow As Integer                            
Dim compareOutlookRow As Integer                            
lastAllScriptsRow = 6                           
compareOutlookRow = 1                           
Range(Cells(1, 1), Cells(lastAllScriptsRow, 1)).Select                          
… code …                            
If IsError(Application.WorksheetFunction.Match(Cells(compareOutlookRow, 2), Selection, 0)) Then                         

```

The search for the first three entries in Column B are successful. When the search for the fourth entry, Bob Carter, is conducted (note the entry of "Bob C Carter" in Column A preventing an exact match), I get a "Run-time error '1004': Unable to get the Match property of the WorksheetFunction class." I get the same error when I use Application.WorksheetFunction.IsNA instead of IsError and when I use the positive approach using IsNumeric instead of IsError. Any help is greatly appreciated.

----

Use `Application.Match` instead of `Application.WorksheetFunction.Match`. The former will _return_ an error that you can trap with `IsError` while the latter throws a run time error which is messier to deal with.

Note that Intellisense does not know that `Application.Match` exists. It does though.

[This is a nice reference on the difference between the two.](http://dailydoseofexcel.com/archives/2004/09/24/the-worksheetfunction-method/)
