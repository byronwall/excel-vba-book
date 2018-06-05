### references to Worksheets

The process for working with Worksheets is the same as all the other Excel Object Model objects: obtain a reference to the object and access it properties.  For the Worksheet, there are a handful of ways to obtain a reference to a Worksheet.  Those include:

* ActiveSheet
* Worksheets(index) or Sheets(index) global objects
* Workbook.Sheets(index) or Workbook.Worksheets(index) with a workbook object
* VBA references (Sheet1, Sheet2)
* Store the reference after creating a new sheet
* Iterating through Worksheets and picking with some criteria
* Copy a Worksheet and then search for the result (see notes below; TODO: add notes)

The basic dividing line of the methods above is when you want to access the Worksheet and what you potentially know about it.  The simplest approach is when you want some code to run on the ActiveSheet because you can just ask for it.  Technically, you can avoid most refereces to the ActiveSheet use the unqualified global references, but this can be lead to errors later.  The business of obtaining a reference to a Worksheet using the other means typically only comes up when you are working with multiple Worksheets.  This is quite common to do.

Once you start workign wtih multiple Worksheets, there are a couple of common things you may want to do:

* Apply the same action to multiple sheets
* Process some data on one sheet based on the data on another sheet
* Move data from one sheet to another
* Move a chart or other object from one sheet to another
* Create a throwaway Worksheet with some information about the rest of your Workbook (e.g. output all the sheet names)

In some of those instances, you are working with multiple sheets becuase you want to do something (e.g. print layout or formatting) to multiple sheets.  In others, you are working on multiple sheets because you know in advance that some task will use data from multiple Worksheets.  For the former case, you are likely to throw you code into a loop across all Worksheets and then use some logic to determine whether or not to apply the action.  In the other case, you will likely use a sheet name or index to directly access the sheet you want.

It is worth mentioning that every Workbook has built in dedicated referneces to the Worksheets which can be used.  These exist as a part of the Object Model.  By default they are called `Sheet1`, `Sheet2`, etc.  These objects are always available and provide a direct refernece to the worksheet.  They can be quite helpful if you rename them from the default names.  A couple of important items about these objects:

* They only exist as objects in the current Workbook.  That is, if you want to access a Worksheet in another Workbook, this approach will not work.  You can technically add a reference to the other Workbook, but I don't recommend doing that.
* Their naming is independent of the actual sheet name displayed in Excel. This cna be incredibly confusing for a new developer (especially if they are not using `Option Explicit`).
* It is very difficult to use these objects to perform some action to multiple Worksheets.

For what it's worth, I've never used the objects directly.  I find myself using the sheet name directly when needed.  This leads to issues with the name being changed, but at some point searching for the string in code is easier than trying to rename the object in the VBE sidebar.  All of the references will break either way.
