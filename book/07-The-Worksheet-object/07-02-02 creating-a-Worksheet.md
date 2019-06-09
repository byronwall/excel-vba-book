### creating a Worksheet

Aside from referencing an existing Worksheet often times the core task of some automtion is to create a new Worksheet. There are a number of reasons you might want to do this:

- A blank sheet is a great starting part for storing some intermediate or final result. It is nearly guaranteed to be the same every time you call for one which is much better than putting new data in an existing sheet.
- You need a blank sheet for the output of some process that is run over a number of items (each analysis gets a new sheet).
- Copying an existing Worksheet and then applying some transformation to the result.
- You created a new Workbook. This adds an extra step but leaves you with the same result as a new sheet alone (unless it was created from a template).

From my own experience, I find that creating a new Worksheet is an absolutely critical task. Very often the goal of using VBA is to automate some task over a range of inputs or possible outputs. This often means that the output for a given command may need to be produced several times. In this case, I regularly create new Worksheets instead of managing the multiple sets of data in one sheet.

In other cases, you may use a temporary intermediate new Worksheet to provide a dumping place for some calculations or other work. This is a much safer approach than to use the existing Worksheet for temporary efforts. Unless you are certain of the contents of an existing Worksheet, there is little reason to avoid creating a new one.

It's worht noting that Excel is quite performant even with a large number of Worksheets. This is especially true if the Worksheets are not linked or related via calculations. My strongest advice on this front is to liberally create new Worksheets and deal with the aftermath later. If you are building a complicated workflow, sometimes the best output is one that is useful but completely disposable. This means that the output is impressive but due to the speed of the automation there is little reason to save or otherwise consume the resulting file. When this is the case, there is no penalty for disorganized Worksheets if the intended product is still there. Let Excel deal with the references and Ranges etc. while you deal with maintainign the rereferences in VBA.

Having said all of that, creating a Worksheet is incredibly simple `Workbook.Sheets.Add()`. That Function will return the Worksheet object which is a reference to the new sheet. The new sheet will have a default name. THe parameters to `Add` control the location of the new sheet with respect to others. It is very, very unlikely that you will create a new Worksheet and not immediately want the sheet reference. That is, you will probably always call `Add` with a preceding `Set` to save the reference. This reference can be as good as gold in an automated workflow since an empty Worksheet is a very powerful starting (and possibly daunting) point.

If you need a copy of an existing Worksheet instead of a blank one, the command is quite simple: `Worksheet.Copy()`. This will create a Copy with parameters for location (TODO: is that true?). The major downside of using `Copy` is that it will NOT return a reference to the newly created Worksheet. This is a real travesty because it means you then have to turn around and do some work to find the newly create Worksheet. My preferred approach is to Copy the Worksheet to the first or last location in the Sheet order and then find it there. Once found, you can move the Worksheet to a desired location and then use the reference.
