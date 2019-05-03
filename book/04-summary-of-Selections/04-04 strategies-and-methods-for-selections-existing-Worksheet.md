## strategies and methods for selections, existing Worksheet

When working with data in an existing `Worksheet`, the main goal is to find the section of the data that you actually want to process. This task can range from trivial to the bulk of the VBA code. A rough overview, starting with trivial is:

- Use the selection -- `Selection`
- Use the ActiveCell -- `ActiveCell` (see later for why these are different)
- Hard-code the address of a single cell -- `Range("A1")` or `Cells(1,1)` (please don't use the latter)
- Name a cell and use that name directly -- `Range("CellName")`
- Iterate through all cells -- `Cells`, `UsedRange`
- While iterating through cells, use some logic to identify if a `Range` is the one you want:
  _ Check the `Value` of the cell
  _ Check if the cell has some property (e.g. `HasFormula`, `HasArray`, etc.) \* Check the `Style` of the cell
- Take an existing `Range`, possibly all cells, and pare it down using:
  _ Move from a known cell to a new spot -- `Offset()`, `End()`
  _ Take a subset of an existing `Range` -- `Cells`, `Rows`, `Columns`, `Areas`
  _ Take a an existing `Range` and change its size -- `Resize()`
  _ Take a super set of an existing `Range` -- `EntireColumn`, `EntireRow`, `CurrentRegion`, `CurrentArray` \* Allow Excel to filter the `Range` based on things it tracks (e.g. value, blank, hidden, etc.) -- `SpecialCells()`
- Identify several `Ranges` and combine them -- `Union()`
- Identify several `Ranges` and use only the common cells -- `Intersect()`
- Pull the `Range` reference from some other object
- Name a cell and use that name indirectly -- `Names("CellName")`
- Ask the user to select the `Range` to use
- Use a function to get a reference -- `Application.Index`
- Search for the cell based on its function or value -- `Find()`
- Process a formula to determine the `Range` it depends on

In addition to those "simple" techniques above, there are more advanced techniques available. Those advanced techniques all rely on some combination of the above options, along with additional logic to manipulate the `Worksheet`. A couple of combination techniques would include:

- Use the Offset-Intersect technique to get a block of data without its header
- Use the `AutoFilter` to filter a data set and then get the visible cells with `SpecialCells()`
- Use one of the techniques above to get a `Range` on one `Worksheet`; grab the corresponding `Range` on a another `Worksheet` to do some processing
