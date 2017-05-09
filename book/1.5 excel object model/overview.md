# overview of Excel Object Model

The Excel object model should focus on the relationship of classes along with the hierarchy.

From the top

TODO: finish this list

This list should focus on the most commonly used parts of the object model.
Add links to other sections of the book with this overview.

Application
  Workbooks -> Workbook
    Worksheets -> Worksheet
      Range -> Range
        Formula
        Value
        Address
        [fomratting things]
      Cells -> Range
      ChartObjects -> ChartObject
        Chart
          Series
            SeriesFormula
          Axes -> Axis
          ChartArea
          PlotArea
      Shapes -> Shape
    Names -> Name
      RefersToRange -> Range

The object model is much easier to work through when declaring variables correctly.  There are a handful of spots (especially with arrays/collections) where the returned object is not helped by Intellisense.

