### Hard-code a cell reference

The second most common way of getting access to a `Range` is to simply give Excel the address of the `Range` to work with. This is a convenient way of working with `Ranges` because it can be easily checked against normal Excel formulas and addresses. The common ways of doing this are using the `Range` and `Cells` functions with the appropriate parameters.

When working with these functions, it is possible to use them "bare" or unqualified. That is, you can just type `Range()` or `Cells()` and it will work. Specifically, it will work on the `ActiveSheet` of the `ActiveWorkbook`. This can lead to some difficulties when working with multiple `Worksheets` or `Workbooks`. If you are working across contexts (`Worksheets` or `Workbooks`), you should generally qualify your reference to the widest context required. This is done by calling the appropriate function on the appropriate object/context. If you have multiple `Worksheets`, you would call `Worksheet.Range()` or specifically `Sheets("SheetName").Range()` in order to access a `Range` on that specific `Worksheet`. If you are working with multiple `Workbooks`, you still only need a reference to the `Worksheet`, but you will have to go through the correct `Workbook` first. This looks like: `Workbooks(1).Worksheets(1).Range`. If you've previously stored a reference to a `Worksheet`, you do not have to use the `Workbook` also; it is very common when working across `Workbooks` to store a `Worksheet` reference as you go (for this reason).

This caveat about qualifying a reference brings up an important point: a `Range` can only refer to cells that are on the same `Worksheet`. You are not allowed to create a `Range` across multiple `Worksheets`. (TODO: what happens if you try this?). If you want to work with `Ranges` on multiple `Worksheets`, you will need to iterate through the `Worksheets`.

#### `Range()`

The `Range()` function is the powerhouse of cell referencing. It works hard to take whatever you give it and return a valid cell reference. It can process the same commands as the address bar in Excel. That is, it will parse:

- a cell reference (`A1`)
- a multi-cell reference (`A1:B5`)
- a discontinuous reference using a union (`(A1, B1, C1)`)
- a discontinuous reference using an intersect (`(A:A 1:1)`) -- Note this will return the cell `A1` which is at the intersection of the two given references. Also note that this way of referencing cells is incredibly rare (I've never used it in a real application).
- a named range (`some_named_range`)
- any application of the multi cell references with named ranges

TODO: can the Range handle a function in it?

Alongside that power of the `Range()`, you can also use it to refer to a group of cells using the corners of the `Range`. This can be used to either return a group of cells in the same row/column, or it can be used to grab a block of data. You are free to give the cells in whatever order you'd like (not required to be top left and bottom right).

This multi-cell version of the `Range()` function is quite powerful when you know or can determine the corners of the `Range` you want. In particular, this works well with the `End()` and `Offset()` functions to build `Ranges` from a single starting point.

If you thought the `Range()` couldn't get any better, it has one last trick up its sleeve. It can also take parameters that are of the `Range` type when building a multi cell `Range`. This is quite powerful because it means you can use any of the techniques to find a `Range` and then get a block of data by feeding them to the `Range()` function. This saves the hassle of calling `Range(someRange.Address, someOtherRange.Address)` just to build the block.

There is one approach to using the `Range` function that is effective but can be a bad crutch. It involves building a `String` to feed to the `Range()` function. This usually looks like `Range("A" & Cells(1,1).Column)` or something similar. There are legitimate cases where this is a quick and easy way out of a problem. It generally involves knowing that you want a cell from a specific row or column while also knowing the other piece (column or row) from an existing cell. You can quickly combine the two to get your reference. There is nothing wrong with building a `String` here, but it might be a sign that there was a better way to get the reference from the start. It can be helpful when working with far to the right columns that are not easily thought of as a number; what column is `AB6` again?

When considering whether and how to use the `Range()` function, the main things to consider are:

- How stable does this code need to be?
- How likely am I to change the address of the cell I want?
- Will a given cell always be in the same place?
- Will a given name always exist?

This questions are pointing to some of the downfalls of `Range()`. The biggest downfall is that if you are going to use `Range("A1")` to refer to cell `A1`, your VBA code will not work if that cell moves for some reason. Furthermore, it can be a real pain to identify when code is failing because of a bad cell reference. I've had it happen countless times now where I hard-code a cell reference, use that in VBA, and then break things completely by adding a row or column somewhere. This is akin to using `VLOOKUP` and inserting a column in the middle of the lookup range; your code will not know or adjust to the new reference. Even worse, depending on what your code does, it's entirely likely that it will run just fine with the mistake. This is the most pernicious type of error to debug in a complicated program.

The upside of this dilemma is that you can quickly remedy the situation by using a named range to refer to the cell. If you name the cell on the Excel side of things, you get the benefit of Excel moving the reference around if the underlying cell moves. This is an incredibly powerful technique. More emphatically, this is the fastest way to "level up" your VBA if you are just getting started. Robust VBA generally relies on named ranges on the underlying spreadsheet. It takes very regular spreadsheets to get away hard-coded references. As a tip, the second time you manually increment 10+ `Range("A1")` calls because of a new row is the last time you want to do that.

A common technique for building macros quickly is to start with hard coded references and convert them to named ranges once the spreadsheet takes form. There is nothing wrong with naming ranges early and not needing them, but it can take more time than it's worth to name the ranges instead of hard-coding a reference. Again, this can burn you quickly if you have to manually change several of those references.

#### Cells()

A convenient but less powerful version of `Range()` is the `Cells()` function. `Cells()` is much simpler since it only requires a row or column number for the reference. This can be useful to quickly grab a reference if you know the row or column number (or both). It's far more likely that you know the Excel reference you want -- `A1` -- than that you know the exact row and column number. It's the column number that is always a pain to determine. Some folks try to get around this by using the `Asc() - 65` approach to get the number for the letter and send that into `Cell1()`. Once you know about the `Range()` function, you'll never touch that madness again.

So, if the `Range()` function is typically more useful and powerful than `Cells()`, why would you ever use `Cells()`? Well, `Cells()` is the entry point for iterating through the cells in a multi-cell `Range`. This use of `Cells` will be covered later on, but it's mentioned here because it's incredibly powerful in that context. Specifically, if you have a `Range` already, you can use `Range.Cells()` to grab a cell within that `Range` at the specific spot. In this way, `Cells()` is actually useful because the indices are smaller and typically correspond to the actual application at hand. Again, this is covered later.

TODO: add a link to the section where iteration is covered
