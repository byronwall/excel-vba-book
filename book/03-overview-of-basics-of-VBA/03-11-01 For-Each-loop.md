### For Each loop

It is not traditional to start with the For Each instead of the For loop, but I personally use the For Each far more so I'll start there.

The For Each loop is used whenever you have an utterable collection. An utterable collection can come from either the Excel object model or your own code. In general, most of the Excel object model returns an utterable collection. This is especially true for Ranges.

TODO: add a list of utterable collections that can be used here

You are not required to put the variable name in the Next line. I recommend not including the variable unless you have tons of code in the loop and are nesting loops. Typically you will rename the variable and then get a compile time error because the variable names don't match. I've never found the variable name in the Next line to help much.

TODO: add an example of a For Each loop
