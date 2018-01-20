### finding a `Range` by paring down (or up) an existing Range

One of the key ways to access a `Range` is to use an existing `Range` and modify it slightly.  This might prompt the question: how do I get the first `Range` in order to use that?  Well, check the previous section for the most common techniques.  You can always start with `ActiveCell` if you just want to see these in action.

Using a `Range` to get the next `Range` really is the bread and butter of serious VBA development. It is a very common pattern to identify a single `Range` in a `Worksheet` that is critical to the rest of the spreadsheet and use that as an "anchor" to access the rest of the cells.  This is particularly common when the data is structured in some way that can be utilized.

When using these techniques, there are a couple of common strategies.  They work by either paring down the current `Range`, moving the current `Range`, or using the current `Range` as the start of some expansion.  Of course, since a `Range` can be used to access a `Range`, you will quickly find yourself chaining these functions together.  That is the true power of these techniques.  Very often you will use 2 or 3 to take a single cell, move to a new spot, resize to cover all of the data and then move over a column to do something.

* Take an existing `Range`, possibly all cells, and pare it down using:
      * Move from a known cell to a new spot -- `Offset()`, `End()`
      * Take a subset of an existing `Range` -- `Cells`, `Rows`, `Columns`, `Areas`
      * Take a an existing `Range` and change its size -- `Resize()`
      * Take a super set of an existing `Range` -- `EntireColumn`, `EntireRow`, `CurrentRegion`, `CurrentArray`
      * Allow Excel to filter the `Range` based on things it tracks (e.g. value, blank, hidden, etc.) -- `SpecialCells()`

#### move to a new spot, `Offset()` and `End()`

There are two simple ways to "move" from a given `Range` to a new `Range`, namely using `Offset()` and `End()`.  Both of these take an existing `Range` and return a new one.  `Offset()` will not modify the size of the current `Range`; it will just move it.  `End()` will always return a single cell even if the starting `Range` was multi-cell.

##### `Offset()`

`Offset(rows, columns)` works by moving the given `Range` over by the parameters given to it.  The nice thing about `Offset()` is that the parameters can be negative to move backwards.  There are a couple of simple use cases for `Offset()`:

* Work your way down or across a group of cells, by `Offsetting()` and setting a reference to the new cell
      * This is often paired with a `While` loop to work down a `Range`
      * This is also helpful when you are not exactly sure what `Range` you want (maybe it's dependent on cell values) so you can't simply assign the correct multi-cell `Range` at the start.
* Use an existing `Range` to get the starting point for a `Range` and move over to a neighbor cell or a blank area to do something
      * THis is common when using one cell's value to determine the value of the next one (e.g. splitting on a delimiter)
      * THis is also common when adding formulas to a spreadsheet.  Find the current data, `Offset()` over a column and apply the formula to all cells.
      * Also helpful when you "just know" that a desired `Range` is some distance away from the `Range` you've got.  This is not the most elegant code at times (since it breaks easily), but it works reliably when you control the spreadsheet.

TODO: add a while loop example
TODO: add a formula example

#### `End()`

`End(xlDirection)` is a powerful function for its specific use case.  It replicates the functionality of the `CTRL+Arrow` keyboard shortcuts.  It will move from the current `Range` as far as possible in a given direction so long as the cells are contiguous. Contiguous in this sense refers to the fact that the cells must not have a blank cell in between them.  A blank cell is any cell that does not have a value _or_ a formula.  The formula part is important because you can use a formula to return `""` while still counting as a contiguous `Range`.

`End()` takes a parameter which is the direction to travel in.  You can go all 4 directions, up/down and left/right.

`End()` will always return a single cell as the reference.  This often means that `End()` is used alongside a `Range(Range, Range)` to get a multi-cell `Range` that spans from the start cell to the end cell.  This is so common of a pattern, that I typically add a UDF that handles this logic directly.

TODO: add the function that is used `RangeEnd`

There are a few common patterns when working with `End()`:

* Use a `Range` that you know is at the top of a block of data and use `End(xlDown)` to get to the bottom of the column.
      * This can be combined with `Range(Range, Range)` to get the full multi-cell `Range` to work through
      * THis technique is very powerful when redefining the `Ranges` of a chart to include all of the cells (this can also be used for formulas too).
* If you know your data has blanks, you can use `End()` to jump to the next non-blank cell.
      * This is helpful if you are trying to fill in blank cells (TODO: add the Waterfall fill here)

##### RangeEnd.md

```vb
Public Function RangeEnd(ByVal rangeBegin As Range, ByVal firstDirection As XlDirection, Optional ByVal secondDirection As XlDirection = -1) As Range

    If secondDirection = -1 Then
        Set RangeEnd = Range(rangeBegin, rangeBegin.End(firstDirection))
    Else
        Set RangeEnd = Range(rangeBegin, rangeBegin.End(firstDirection).End(secondDirection))
    End If
End Function
```

##### RangeEnd_Boundary.md

```vb
Public Function RangeEnd_Boundary(ByVal rangeBegin As Range, ByVal firstDirection As XlDirection, Optional ByVal secondDirection As XlDirection = -1) As Range

    If secondDirection = -1 Then
        Set RangeEnd_Boundary = Intersect(Range(rangeBegin, rangeBegin.End(firstDirection)), rangeBegin.CurrentRegion)
    Else
        Set RangeEnd_Boundary = Intersect(Range(rangeBegin, rangeBegin.End(firstDirection).End(secondDirection)), rangeBegin.CurrentRegion)
    End If
End Function
```

#### Take a subset of an existing `Range` -- `Cells`, `Rows`, `Columns`, `Areas`

The subset functions work by providing you with a Range that is created from another Range based on some condition.  They can be quite useful for building a workflow that makes it very explicit how you are trying ot iterate through a Range or waht you are searching for.  The idea is that you know your starting Range contains some pieces that you would like to tierate through.  The grouping goes from smallest unit to largest:

* Cells will return a "flat" list of all cells with in the Range.  No grouping left.
* Rows and Colujmns will each return a new iterable object built of the previous Range sliced into its Rows or Columns.  If call them in order, it will look the same as iterating through Cells except that the order may be difference (TODO: how does this work?).  Be sure that if oyu want to yuse htese, avoid the properties with the "s".  If you call Row ro Column, you will just get a number instead of a group of Ranges
* Areas will return a group of cells that may contain groups of Rows or Columns or just individual Cells.  Areas are commonly built by users using `CTRL` to select multiple things or by VBA which uses `Union` to build Ranges.

TODO: add some specific code related to Columns and Rows

TODO: give an example of using Areas

#### Take a an existing `Range` and change its size -- `Resize()`

`Resize()` is a straightforward function that does exactly what you expect.  It takes a current `Range` and resizes it to contain the number of rows and columns specified.  The most common uses of a `Resize()` are:

* You know where you want some output to start and its size, so you `Resize()` to get a `Range` that will hold all of the data.
* You know that some data starts at a given cell and its size, so you `Resize()` and call `Value` to get an array of that data.
* You would like to extend or change a formula based on some condition, so you `Resize` and apply the formula down the line

In general, these uses follow a pattern: you know what size you want the `Range` to be (or can compute the size) and `Resize` gives you the `Range` back.  This is one of the least controversial of the `Range` methods.  Enough said.

TODO: how does this handle negative numbers
TODO: how does this handle a multi-cell range, does it always pick top left?

#### Take a super set of an existing `Range` -- `EntireColumn`, `EntireRow`, `CurrentRegion`, `CurrentArray`

These "super set" functions work by taking a starting point and expanding it to include more cells.  These will grow the `Range`.  Of the four listed above, `CurrentArray` is the only one that requires some special case.  That is, the current cell must be a part of an array formula.  The others will always work.  These functions are best thought of with their keyboard shortcut equivalents:

TODO: extract this table along with others and make a single big table somewhere

shortcut | Range function
--- | ---
SHIFT + SPACE | `EntireRow()`
CTRL + SPACE | `EntireColumn()`
CTRL + A | `CurrentRegion()`
CTRL + / | `CurrentArray()`

`CurrentRegion` is really only as useful as the data on the spreadsheet.  If you have a large block of data, it works well to get the entire region.  If you have blanks in your data, it's a bit of an unknown to know in advance what `CurrentRegion` will give you.  Typically, if you know you have a block of data, it can be a quick shortcut to using `End()` twice.  In general, I avoid it.

`EntireRow` and `EntireColumn` are somewhat special because they can be used to make modifications to the rows and columns in Excel.  In particular, they are needed if you want to insert a row/column, delete a row/column, change the row/column formatting, or change the height/width of the row/column.  You can also use `Range("A:A")` or similar ot get a reference to the entire column, but it is much simpler to have a reference to a `Range` of a single cell and work out from there.  Even better, if you have a multi-cell `Range`, the `Entire` functions will return the combination of all the rows or columns contained in the `Range`.

In addition to modifying the rows/columns of a `Worksheet`, the `Entire` functions also work very nicely with `Intersect()` to get group of cells that are in a specific row/column.  The `Entire` functions are generally much nicer than trying to build the `Range` from address or any other technique.

TODO: is this true?  Does it work for a multi-cell in this way?

#### Allow Excel to filter the `Range` based on things it tracks (e.g. value, blank, hidden, etc.) -- `SpecialCells()`

The final function in this round up is also the most powerful at times: `SpecialCells()`.  This function works by taking a parameter how which "special" cells to return.  Special is a bad name here, because the most common uses of `SpecialCells` are to grab cells that are formula, values, blanks, or visible.  These are some of the more mundane properties of a cell.  Name aside, `SpecialCells()` can really take your VBA to the next level with very little effort.

An example: if you have ever iterated through `UsedRange` or `Cells` with something that checks for `rng.Value = ""` then you could have saved a loop by using `SpecialCells(xlCellTypeBlanks)` instead.  This will return a new `Range` that only contains the blank cells.  There are similar special types for other things that commonly come up.

One particular application of `SpecialCells` is when working with the `AutoFilter` which will cause rows to be `Hidden`.  You can get a `Range` that contains all of the visible rows which is the same as the rows which satisfy the filter.  If your data is well structured or can be filtered, this ends up being a great way to push the burden of filtering onto Excel instead of having all that logic in VBA.

You can also use `SpecialCells` to quickly return a list of those cells which have a value (or formula) if you have a large block of sparse data.  Once you have all of those cells, you can `Intersect()` the `EntireColumn` (or row) with the header of the data.  This allows you to move quickly through data without having ot build addresses or remember where specific things are.  In general, this highlights an important strategy: if you can obtain `Ranges` with the areas that are critical, you cna quickly manipulate those `Ranges` to perform some action.  You can spend less time building finding cells and `Ranges` once you know how to work and combine these functions.

TODO: add the table manipulation code here to give an example of that
TODO: consider adding an example of using SpecialCells with filtering
