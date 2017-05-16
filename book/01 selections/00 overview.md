# summary of Selections

The selections chapter needs to focus on the ways to access Ranges from VBA.      These should cover the various ways to mimic normal ways of selecting along with VBA special stuff.

## ways to get a Range object

* calling `Range`
* calling `Cells`
* from an existing Range
      * Cells
      * Rows
      * Columns
      * SpecialCells
      * Offset
      * Resize
      * EntireRegion
      * End
* from a Name object
* using `Selection`
* using `ActiveCell`
* using `Union` and `Intersect`
* using `Find`
* using `UsedRange`
* using `Application.Index` (or is it only WorksheetFunctions?)
* using `CurrentArray`

## some common patterns combining these techniques

* the Offset-Intersect approach (move a block down and intersect with the original)
* the Offset-Resize pattern when you move to a cell and expand the selection based on something
* the Union-Delete approach to getting a Range to delete
