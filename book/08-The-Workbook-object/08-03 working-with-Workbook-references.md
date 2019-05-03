## working with Workbook references

There are a couple of ways to obtain a refernece to a Workbook that are useful:

- ActiveWorkbook - refers to the Workbook that has focus
- ThisWorkbook - refers to the Workbook which contains the code that is executing
- Workbooks.Open() - will open a Workbook and return a reference
- Workbooks(index) - will grab a refernece to the currently opened Workbook
- Workbooks.Add() - will create a new blank Workbook or a Workbook according to a supplied template

I find that all of those approaches are used equally across my code. The one exception might be ThisWorkbook which I typically avoid. In reality, I should probably use it more becasue I find myself going to some length to maintain a reference to a Workbook while opening or creating Workbooks.

For Workbooks, the biggest thing to be aware of that there are a number of unqualified references that exist within VBA that are a part of the ActiveWorkbook. Those include:

- Worksheets and Sheets
- Names?

These unqualified referecnes can really bite you when you are expecting it. The problem with unqalified references is taht they work great initially, before the workflow becomes complex. They will then silently fail later when you start creating new Workbooks and otherwise changing the focus or active Workbook. The problem is that nearly all of the unqualified references apply to the ActiveWorkbook. Working with Workbooks is the one task that will often change the focus of Excel regardless of how you create things.
