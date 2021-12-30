## useful properties of the Workbook

Although I have railed against the Workbook object, there are a handful of things that it can do:

- Reference `Names` which contains all of the global named ranges
- Others?
- Charts?

### Worksheets vs. Sheets

WHen working with Worksheets, there are a pair of objects which will provide access to the underlying Sheets. They are different in how they handle Charts which are visible as a Worksheet. The rule is: Sheets will return the Charts, whereas Worksheets will only return the list of objects which are actually Worksheets. If you do not use Charts as Worksheets, then you will never notice a difference between these two objects. The one thing you will notice is that the ActiveWorksheet will not be of type Worksheet which means that you can never get Intellisense on one of the most useful objects.
