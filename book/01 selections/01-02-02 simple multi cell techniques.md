### some simple techniques for finding a multi-cell `Range`

The simple selection technique for working with multiple cells consist of:

* Iterate through all cells -- `Cells`, `UsedRange`
* Building a range from the corners -- `Range()`

The previous section identified the simplest techniques for obtaining a reference to a `Range`.  Those techniques touched on single and multi-cell `Ranges`.  There are a couple of additional techniques for obtaining a multi-cell `Range` that are used commonly.

The typical goal of these multi-cell calls is to take the reference and iterate through the cells.  To iterate through the cells, there are two techniques, `For Each` and `For` loops.  The former is vastly preferred to the latter in nearly all cases.  I'll say that again, if you're iterating through cells, you should strongly prefer to use a `For Each` loop instead of a simple `For` loop.  Those two examples look like:

TODO: add code samples for `For` and `For Each` loops

#### `Cells`

The `Cells` call exists on several different objects.  The easiest way to access it is via the bare, unqualified, reference -- just type `Cells`.  It applies to the `ActiveSheet` of the `ActiveWorkbook`.  Typically, you should avoid iterating all `Cells` unless you know you will break out of the loop at some point.  There are a lot of cells in a `Worksheet`, and your code will grind to a halt working through rows 10100 to 132000 doing a bunch of nothing on empty cells.

#### `UsedRange`

`UsedRange` is available on a `Worksheet`.  It also exists as a bare unqualified reference applying to the `ActiveSheet` of the `ActiveWorkbook`.  The `UsedRange` is a slightly complicated function but its goal is to provide you a `Range` that provides a bounding box on all of the used cells in the current `Worksheet`.  The complication of `UsedRange` comes when determining what is a "used" cell.  Excel will consider a cell used if it has a non default property for its value or formatting.  The formatting part of the definition can throw you for a loop because it's possible to change the formatting in a non-obvious way (e.g. it's impossible to spot the font size of an empty cell).  There are several well-regarded folks who will advocate against the `UsedRange` in all cases.  Their argument is that the `UsedRange` is too undependable because it can be thrown off too easily.  In my experience, the `UsedRange` is a powerful way to leverage Excel tracking the internal state of the spreadsheet.  You can also avoid most of the issues with the `UsedRange` not matching expectations by taking care of the state of the spreadsheet.  If a `Worksheet` was under your control, there's no reason to avoid the `UsedRange`.  As a first tip, the `UsedRange` matches the scrollbars around the spreadsheet.  If the scrollbars stop scrolling when you reach the "end of the spreadsheet", then the `UsedRange` is good to go.  You can also do a quick test with `UsedRange.Address` or `UsedRange.CountLarge` to see what it refers to.  Again, I think the arguments against the `UsedRange` are overly cautious, and it's a great command in a well managed spreadsheet.

TODO: is `UsedRange` available bare?
