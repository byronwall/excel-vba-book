### Disabling calculations

The most common approach to controllign calculatiojns is to simply disable them. To "disable" the calculations, is really to set the CalculationMode to Manual. It does not actually disable calculations, but instead it prevents the automatic calculations updates from firing like normal. The spreadsheet still maintains its normal model of calculations; they just don't run. This is an incredibly common approach to speeding up the perofmrance of VBA code. The performance boost results form the fact that when VBA code executes, it is very tihglty coupled to the normal Excel operaitons that take place. When you use VBA to set a `.Value` equal to some new value, it is functioanlyl equivlabet to manually entering the value. Behind the scenes, Excel will fire off the normal Change events and update the dependent cells. THis can become a bottleneck because VBA is able to rapidly fire off `.Value` changes. So rapidly, that processing all of the associated stuff can become a limitation. It is more or less guaranteed that you will run into this issue once you start writing VBA code. It is so common, that you will likely memorize the fix:

TODO: check this code

```vba
Application.CalculationMode = xlManual
Application.ScreenUpdating = False
Application.EnabledEvents = False

Application.EnabledEvents = True
Application.ScreenUpdating = True
Application.CalculationMode = xlAutomatic
```

Why does this code make everything faster? Well, it disables the slowest steps of Excel keeping track of your spreadsheet: visual updates, calculations updates, and other events. Turning all of those off will dramtically remove the bottlenecks to your code. What's the downside? Well, all of that stuff exists for a reaosn and it's possible you need it to kepe funcitojning for some VBA operaitons. The non-calculatojn optiojsn ar ecovered in subsequent chapters, so we'll focus on the claculatiojn part now.

What happens when you disable calcualtiojns? THis is the key concept to understand to make sure your spreadsheets do not break when you go looking for performances. So what changes?

- Dependent cells are not updated. The "chain" is processed to its end on every update. Note, updates are sent downstream. Not all cells are updated, unless your Workbook contains a VOLATILE function.
- Charts and other functional graphics do not update. Internally, they don't change at all. It's not just a matter of the visuals being hiddne, they are not calculated.
- Less important items:
  - Conditional formatting will not update.

So why might those things matter? The biggest reason is that if your VBA code depends on the state of the spreadsheets, then you are likely depending on calculations at some point. This means that you need to split you rcode into semgents where you are not worried about cell values and those wehre you are. An example:

You are building a tool to process data from a CSV file. You have been told that you should delete data that is in the 0th to 10th percentile of a cost column. Unforateunyl, the data needs to be preprocessed in order to create an accurate cost column. Your CSV file contains a mess of extra text and other issues when need to be removed. Your workflow then is:

1. Import the CSV data
2. Preprocess the cost column to clean up the mess
3. Remove the rows below the 10th percentile.

You do a quick test and have no problem importing the CSV data. You've gone ahead and worked out the preprocessing logic... only took a couple calls to `Split` and `Trim`. YOu also went ahead and added a new colujmn to compute the `PERCENTILE` based on the now cleaned result. This is looking great on your 100 row test data set. Your set your application loose on the 90,000 "real" data and quickly find that it will not complete within 10 minutes. What's going on here? THe most likely provlem is that your new PERCENTILE column is being reculated every time a preprocessed data cell is being added back to the spreadsheet. Your processing code looks like:

```vba
For Each rngCell in rngData
    rngCell.Value = CleanUpThisMess(rngCell)
Next
```

If `rngData` contains 90,000 cells, then your update code will call for at least 90,000 full Worksheet recalculations. Even worse, your PERCENTILE formula requires the entire colujmn of data and so all cells have to update every time. `90,000 x 90,000` quickly becomes a problem.

So, why is the PERCENTILE function updating after every change? Do we really care what the intermediate values are? No.

This is why you want to have control of the calculations. in this case, you kknow that the processing code is not affected by the value of the PERCENTILE column. We only need the static data available in order to complete the processing. The fix here is to turn calcuaitliosn to manual during the processing step so that you do not incur 90,000 extra recalculations.

Once the processing is done, what do we do with the calculatiojn mode? Well, that depends on how we do the deletion. There are a couple of optionsm:

- Turn on an `AutoFilter` and do a FILTER-DELETE to remove all the rows in one shot.
- Iterate through the rows, one by one, and remove those which are in the 10th or lower percentils

Looks like either will work, but hwo does calculatiojn mode affect things? Well, if you go with the latter option, you will find that your PERCENTILES will update after each deletion. This is not the behavior you intended. You somehow want to remember the PERCENTILE value before oyu started the deletions. The solution then is to contorl the calculation mode again. Here, we are controllign things for **accuracy**. Our deletiojn appraihc will not work if we allow cells to update as we go.

Pro tip: if you are deleting cells, you should pretty much never go a row, column, or cell at a time. Instead you should build a `Range` of cells to be deleted using `Union` and delete them in one shot using `Delete`. This approach is called a `UNION-DELETE` and avoids all of the issues described above. It's also the fatest approach since it does a single deletion.
