## Controlling calculations

When you are creating macro workflows, there are a number of tools at your disposal to control calculations flow. Before describing those tools, it's worth stepping back and discussing why you might want to control the calculation flow. There are a couple of common reason:

- Performance. Your code will run faster if you control the calculation process. This mainly involves disabling automatic calculation at key points.
- Accuracy. For some types of calculations, you need to tightly control the calculation flow for accuracy. This is often the case if you are building a spreadsheet that does some form of recursion or self reference.
- Usability. There are some situations where you are interacting with calculations and need to prevent the normal behavior. The most common is when you add Workbook events like `Change`.
- Profiling. If you are building a code profiler (i.e. a tool that tracks execution time of your code) you must control calculations in order to get the tracking right.

We'll get back to the applications, but it's also worth hitting the high points on how you can control the calculation. THe main knobs:

- Disable application wide
- Disable for a Worksheet
- Manually calculate a Range, Worksheet, or Application

THe types of changes you will make are fairly tightly couple to the applications above. In general, for performances nad usability reasons, you will be disable calculations. For accuracy or profiling applications, you will manually walking the calculation through.

### Disabling calculations

The most common approach to controlling calculations is to simply disable them. To "disable" the calculations, is really to set the CalculationMode to Manual. It does not actually disable calculations, but instead it prevents the automatic calculations updates from firing like normal. The spreadsheet still maintains its normal model of calculations; they just don't run. This is an incredibly common approach to speeding up the performance of VBA code. The performance boost results form the fact that when VBA code executes, it is very tightly coupled to the normal Excel operations that take place. When you use VBA to set a `.Value` equal to some new value, it is functionally equivlabet to manually entering the value. Behind the scenes, Excel will fire off the normal Change events and update the dependent cells. This can become a bottleneck because VBA is able to rapidly fire off `.Value` changes. So rapidly, that processing all of the associated stuff can become a limitation. It is more or less guaranteed that you will run into this issue once you start writing VBA code. It is so common, that you will likely memorize the fix:

TODO: check this code

```vba
Application.CalculationMode = xlManual
Application.ScreenUpdating = False
Application.EnabledEvents = False

Application.EnabledEvents = True
Application.ScreenUpdating = True
Application.CalculationMode = xlAutomatic
```

Why does this code make everything faster? Well, it disables the slowest steps of Excel keeping track of your spreadsheet: visual updates, calculations updates, and other events. Turning all of those off will dramatically remove the bottlenecks to your code. What's the downside? Well, all of that stuff exists for a reason and it's possible you need it to keep functioning for some VBA operations. The non-calculatojn options ar ecovered in subsequent chapters, so we'll focus on the calculation part now.

What happens when you disable calculations? This is the key concept to understand to make sure your spreadsheets do not break when you go looking for performances. So what changes?

- Dependent cells are not updated. The "chain" is processed to its end on every update. Note, updates are sent downstream. Not all cells are updated, unless your Workbook contains a VOLATILE function.
- Charts and other functional graphics do not update. Internally, they don't change at all. It's not just a matter of the visuals being hidden, they are not calculated.
- Less important items:
  - Conditional formatting will not update.

So why might those things matter? The biggest reason is that if your VBA code depends on the state of the spreadsheets, then you are likely depending on calculations at some point. This means that you need to split you rcode into segments where you are not worried about cell values and those where you are. An example:

You are building a tool to process data from a CSV file. You have been told that you should delete data that is in the 0th to 10th percentile of a cost column. Unforateunyl, the data needs to be preprocessed in order to create an accurate cost column. Your CSV file contains a mess of extra text and other issues when need to be removed. Your workflow then is:

1. Import the CSV data
2. Preprocess the cost column to clean up the mess
3. Remove the rows below the 10th percentile.

You do a quick test and have no problem importing the CSV data. You've gone ahead and worked out the preprocessing logic... only took a couple calls to `Split` and `Trim`. You also went ahead and added a new column to compute the `PERCENTILE` based on the now cleaned result. This is looking great on your 100 row test data set. Your set your application loose on the 90,000 "real" data and quickly find that it will not complete within 10 minutes. What's going on here? THe most likely problem is that your new PERCENTILE column is being reculated every time a preprocessed data cell is being added back to the spreadsheet. Your processing code looks like:

```vba
For Each rngCell in rngData
    rngCell.Value = CleanUpThisMess(rngCell)
Next
```

If `rngData` contains 90,000 cells, then your update code will call for at least 90,000 full Worksheet recalculations. Even worse, your PERCENTILE formula requires the entire column of data and so all cells have to update every time. `90,000 x 90,000` quickly becomes a problem.

So, why is the PERCENTILE function updating after every change? Do we really care what the intermediate values are? No.

This is why you want to have control of the calculations. in this case, you know that the processing code is not affected by the value of the PERCENTILE column. We only need the static data available in order to complete the processing. The fix here is to turn calcuaitliosn to manual during the processing step so that you do not incur 90,000 extra recalculations.

Once the processing is done, what do we do with the calculation mode? Well, that depends on how we do the deletion. There are a couple of options:

- Turn on an `AutoFilter` and do a FILTER-DELETE to remove all the rows in one shot.
- Iterate through the rows, one by one, and remove those which are in the 10th or lower percentiles

Looks like either will work, but hwo does calculation mode affect things? Well, if you go with the latter option, you will find that your PERCENTILES will update after each deletion. This is not the behavior you intended. You somehow want to remember the PERCENTILE value before you started the deletions. The solution then is to control the calculation mode again. Here, we are controlling things for **accuracy**. Our deletion approach will not work if we allow cells to update as we go.

Pro tip: if you are deleting cells, you should pretty much never go a row, column, or cell at a time. Instead you should build a `Range` of cells to be deleted using `Union` and delete them in one shot using `Delete`. This approach is called a `UNION-DELETE` and avoids all of the issues described above. It's also the fatest approach since it does a single deletion.
