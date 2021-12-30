## limitations of UDfs

This section will focus on the aspects of UDFs where you are limited. There are couple of key things to remember here:

- A UDF is not allowed to change the Workbook, Worksheet, or a Range -- no side effects are allowed
- A UDF will only update if the cells it refers to change
- You can mark a UDF as Volatile, but this may create other problems (namely speed)
- UDFs are allowed to use global variables but you can wreck this process by having errors while they execute
- UDFs inside an addin can pollute a spreadsheet that might be used by someone without that addin
- You can debug a UDF but not by using the Evaluate Formula option that might be familiar to more people

### no side effects

The biggest temptation of a UDF is one of the few things that is not allowed -- you are not allowed to have a side effect from a UDF. This generally comes up when you want to change something about the Range that the UDF is referring to or being called from. You think: "I'd just love to color this cell red if the UDF detects some state while executing". This thought comes up because it'd be nice to have the UDF update when called and even better if you can avoid dealing with conditional formatting. Alas, this is not allowed. The UDF must execute without making a change to the spreadsheet. This generally makes sense if you think about how Excel goes about calculating the spreadsheet. It makes a map of how cells are related and then proceeds to calculate the values in an order where each cell that depends on another is calculated in the precise order that is required. This process allows Excel to complete as fast as possible, without errors, and while using as many CPU cores as possible. If your UDF is able to change the spreadsheet after Excel has determine the order of calculations, then it becomes impossible to ensure that the spreadsheet is still correct. Because of this, Excel does not allow sde effects from a UDF.

The other aspect of this limitation that comes up often enough in practice is hat you cannot use a Worksheet function that modifies the spreadsheet even if you intend to undo that function. For example: I have attempted to use the AutoFilter inside a UDF in order to determine how many times some condition showed up in a table. This is not allowed even though I intended to undo the AutoFilter before returning from my UDF. This limitation also applies to Copy/Paste and other common functions.

### when does a UDF update

The next limitation to consider is that a UDF will only update when the Ranges it refers to are changed. This is related to the dependency tree described above. Excel will only call your UDF if one of the cells that it directly depends on it is updated. This is important because you have access to the entire Workbook inside a UDF so you can create a situation where your UDF _should_ update something, but it doesn't because it does not know that it should have been updated. This si discussed later, but the quick way around this limit is to mark your UDF as Volatile. See the warnings later related to this.

A common example of when this sort of issue pops up is when you are using a reference to a Range inside the UDF that is computed only inside the UDF. For example, you want to do some statistics for a single Range that are dependent on a larger Range of data. You can write a UDF that takes the single cell as a parameter but then compute the larger Range inside the UDF without having to refer to it. Maybe that larger Range is a mess via normal Excel so you've skipped that step. Well, be aware that your UDF will only calculate for the even cell if the cell it refers to changes. This means that the larger group may change -- and invalidate your current result -- but if the single cell stays the same, then your UDF will not update that cell.

This same issue pops up if you are using properties of the Range that are not a part of the calculation model for Excel. That is, there are some changes which will not trigger a recalculation from Excel. These are typically related to using the formatting of a cell in a UDF. A very common example is returning the Range.Text from a UDF so that you can get the value exactly as it is displayed in the spreadsheet. If you change the format of the cell, you are not guaranteed to have the UDF called updating your UDF value.

### using Application.Volatile

Mentioned above, there is one surefire way to ensure that your UDF will be called whenever there is a change anywhere on the spreadsheet: mark the function as Volatile. This is done by calling Application.Volatile somewhere in your UDF. TODO: is this right? Once you have made this call, your UDF will be called anytime a calculation is done. This also means that anything that depends on your cell will be recalculated every time. There is a huge upside to using Volatile UDFs in certain instance: you are guaranteed that they represent the correct value. THe downside is that your UDF is being called constantly which means that if it is slow, your entire spreadsheet will be slow. If your UDF is littered across 10,000 cells, it will be run 10,000 times even if only a single cell changed. It is easy to underestimated how much this can slow down a Workbook. Having said that, sometimes speed is not a factor and you just want things to eb correct.

There are other functions (INDIRECT and OFFSET are the main ones) in Excel that are volatile, so it is not some awful thing to do necessarily. You should mark something as Volatile however only as a last resort or possibly as a first resort if you're just punching something out.

To avoid using Volatile, you may be able to have your UDF take an additional parameter to ensure that it is on the calculation chain of all the cells it depends on. Note: you don't actually have to use the parameters for anything, but if they appear in the UDF call, it will force Excel's calculation tree. Continuing with the statistic example from above, if you know that all of the data that could change is in columns B and C, you can simple send B:C in as a parameter to the UDF. This ensures that a change in those columns will force the UDF to call. You can then continue to compute the Range using your more complicated logic. This is somewhat wasteful and means you have extra parameters which don't do anything, but it can be a cleaner (and faster) solution than using Volatile.

### beware of global variables

VBA allows you to declare a variable outside of any Sub or Function definition. These are typically called global variables because they can be accessed from any code. This means that you can create some variables in a Sub and then use them in subsequent UDF calls. A good example is loading up a database of information and then using that information inside the UDF. This can be nice because then you do not have to load the data every time you call the UDF. I've used this effectively when doing unit conversions with UDFs.

The downside to this approach is that it seems to be relatively easy to corrupt those global variables if you have errors while the UDF runs. I've had it happen where that loaded database becomes corrupted somehow and then all of the dependent cells start to fail when their UDF is called. This type of error can be quite difficult to track down because it may not be obvious why the variable was corrupted.

### beware of UDFs in addins

A personal addin is a great way to organize helper code without constantly created macro enabled files to use the code. For Subs this works great because there is no lasting trace that a Sub was run, at least in terms of code in the file. For a UDF however, your UDF call will be a part of the spreadsheet. This does not force the spreadsheet to be a macro enabled one -- which is great -- but it does mean that anyone using the spreadsheet needs access to the UDF code. This creates a problem when you get comfortable using UDFs in an addin but then save the workbook with them in there. You have effectively "polluted" the workbook with addin UDF names which may or may not be available to others. This is fine if the addin truly is critical to the workbook, but it can create a mess for others if you're using UDFs for your own help and make a spreadsheet that others cannot use.

The solution to this problem is to simply save the UDF as a Module in the spreadsheet, but this requires you to save the Workbook as macro enabled.

A rule I like to follow is simply: if I know that a UDF is required for the spreadsheet and that UDF is currently in an addin, I force myself to move the code into the Workbook and save as macro enabled. This can be a pain, but it's all too common that a Workbook is saved with a UDF from an addin, that addin changes or becomes unavailable, and now your Workbook is broken. It's best to avoid this scenario especially if you work with others who are not macro savvy.

### debugging UDFs is different

Most folks are familiar with the "Evaluate Function" feature of Excel which will help you walk through a function's evaluation in the order that Excel evaluates things. This can be incredibly helpful for array formulas where it's not always obvious the order Excel will do things in. Your UDF will also be evaluated in that feature, but it will not step through the logic of your UDF. This might seem obvious, but it's worth mentioning. IF you want to debug the logic of your addin, you need to set a breakpoint and actually debug the code. See the later section on this for the details.

TODO: add link to that section
