### some simple techniques for finding a `Range`

The simple selection techniques consist of:

- Use the ActiveCell -- `ActiveCell` (see later for why these are different)
- Use the selection -- `Selection`
- Hard-code the address of a single cell -- `Range("A1")` or `Cells(1,1)` (please don't use the latter)
- Name a cell and use that name directly -- `Range("CellName")`

These are considered simple, but their simplicity means they are commonly used. These techniques can return a `Range` that represents either a single cell or multiple cells or a group of discontinuous cells. The one exception to this is the `ActiveCell`; it is always a single cell.

#### Selection and ActiveCell

The `Selection` and `ActiveCell` commands both work based on what is currently going on with the active spreadsheet. In particular, they work on the current selection of the `ActiveSheet` in the `ActiveWorkbook`. For a normal workflow, the active sheet and workbook are the ones with focus (or that last had focus). When working through an involved workflow, you can control the `ActiveSheet` and `ActiveWorkbook`. In general, you should not use these commands in an involved workflow without a very good reason.

##### Selection

`Selection` is a catch all object that refers to anything that is selected. If the current selection is a group of cells, then you get a `Range`. If instead the selection is a Chart, Shape, button, or some other non-`Range`, then you will get an error if you assume that it has type `Range`. When working with the `Selection`, it is always good to assign a new `Range` variable equal to the `Selection`. This ensures that you get Intellisense for commands and also ensures that VBA will throw an error if the `Selection` is something other than a `Range`.

##### ActiveCell

The `ActiveCell` always refers to a single cell. If the current `Selection` is a single cell, then these will refer to the same `Range`. If the current `Selection` is a multi-cell `Range`, then the `ActiveCell` is the cell that currently has focus. When normally editing cells, you have some control over which cell in a multi-cell `Range` is active. This can be changed by hitting `CTRL+.`, `SHIFT+Enter`. This functionality in Excel is what allows an array formula to be applied to a larger range. You select a multi cell `Range` and then enter the formula with `CTRL+SHIFT+Enter`. This in turn will apply the formula to all cells.

TODO: what happens when the `Selection` is not a `Range`? Does this still work?
