## creating and managing Worksheets

This section will focus on how to create a Worksheet and get a reference to new Worksheets.  In addition to that, it will discuss managing Worksheets, including rearranging and deleting them.

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

### creating a Worksheet

Aside from referencing an existing Worksheet often times the core task of some automtion is to create a new Worksheet.  There are a number of reasons you might want to do this:

* A blank sheet is a great starting part for storing some intermediate or final result.  It is nearly gauranteed to be the same every time you call for oen which is much better than putting new data in an existing sheet.
* You need a blank sheet for the output of some process that is run over a number of items (each analysis gets a new sheet).
* Copying an existing Worksheet and then applying some transformation to the result.
* You created a new Workbook.  This adds an extra step but leaves you with the same result as a new sheet alone (unless it was created from a template).

From my own expereince, I find that creating a new Worksheet is an absolutely critical task.  Very often the goal of usign VBA is to automate some task over a range of inputs or possible outputs.  This often means that the outut for a given command may need to be produced several times.  In this case, I regularly create new Worksheets instead of managing the multiple sets of data in one sheet.

In other cases, you may use a temporary intermediate new Worksheet to provide a dumping place for some calculations or other work.  This is a much safer approach than to use the existing Worksheet for temporary efforts.  Unless you are certain of the contents of an existing Worksheet, there is little reason to avoid creating a new one.

It's worht noting that Excel is quite performant even with a large number of Worksheets.  This is especially true if the Worksheets are not linked or related via calculations.  My strongest advice on this front is to liberally create new Worksheets and deal with the aftermath later.  If you are building a complicated workflow, sometimes the best output is one that is useful but completely disposable.  This means that the output is impressive but due to the speed of the automation there is little reason to save or otherwise consume the resulting file.  When this is the case, there is no penalty for disorganized Worksheets if the intended product is still there.  Let Excel deal with the references and Ranges etc. while you deal with maintainign the rereferences in VBA.

Having said all of that, creating a Worksheet is incredibly simple `Workbook.Sheets.Add()`.  That Function will return the Worksheet object which is a refernece to the new sheet.  The new sheet will have a default name.  THe parameters to `Add` control the location of the new sheet with respect to others.  It is very, very unlikely that you will create a new Worksheet and not immediately want the sheet refernece.  That is, you will probably always call `Add` with a preceding `Set` to save the reference.  This reference can be as good as gold in an automated workflow since an empty Worksheet is a very powerful starting (and possibly daunting) point.

If you need a copy of an existing Worksheet instead of a blank one, the command is quite simple: `Worksheet.Copy()`.  This will create a Copy with parameters for lcoation (TODO: is that true?). The major downside of using `Copy` is that it will NOT return a reference to the newly created Worksheet. This is a real travesty because it means you then have to turn around and do some work to find the newly create Worksheet.  My preferred approach is to Copy the Worksheet to the first or last location in the Sheet order and then find it there.  Once found, you can move the Worksheet to a desired location and then use the reference.

### removing a Worksheet

If you need to delete a Worksheet, it is a simple command again: `Worksheet.Delete`.  The one downside to this command is that it will fire off a warning prompt if the Worksheet contained any data or was otherwise not "blank".  This warning box will stall the execution of your VBA until it is addressed.  This is a major issue for any serious workflow since your users will have to constantly click "Yes" to delete the Worksheet but they may also have no idea what they are deleting.  To avoid this issue, you will nearly ALWAYS wrap the `Delete` command with the comannds to disbale and then reenable the alerts.  The typical code looks like:

```vb
Application.DisplayAlerts = False
Worksheet.Delete
Application.DisplayAlerts = True
```

When doing this dance, be absolutely certain that you reenable the alerts.  Excel will not do it for you.  You may benefit from creating a new helper Sub which contains the above code as a `DeleteSheet` command to avoid constantly adding those alerts.

TODO: add a note about when to create a new Worksheet vs. a new Workbook and the pros/cons there (maybe put this in the workflow section of book)

### rearranging Worksheets

To rearrange the Worksheets, the command is simple: `Worksheet.Move(Before, After)`.  The parametrs there will indicate the sheet ot place it before or after. The real task here is determining whcih sheet to refernece there, but finding that reference is the same task that is described up at the top of the seciton.

#### AscendSheets.md

TODO: move the AscendSheets code elsewhere or delete (not helpful here)

```vb
Public Sub AscendSheets()

    Application.ScreenUpdating = False
    Dim targetWorkbook As Workbook
    Set targetWorkbook = ActiveWorkbook

    Dim countOfSheets As Long
    countOfSheets = targetWorkbook.Sheets.Count

    Dim i As Long
    Dim j As Long

    With targetWorkbook
        For j = 1 To countOfSheets
            For i = 1 To countOfSheets - 1
                If UCase(.Sheets(i).name) > UCase(.Sheets(i + 1).name) Then .Sheets(i).Move after:=.Sheets(i + 1)
            Next i
        Next j
    End With

    Application.ScreenUpdating = True
End Sub
```