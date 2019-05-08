### no side effects

The biggest temptation of a UDF is one of the few things that is not allowed -- you are not allowed to have a side effect from a UDF. This generally comes up when you want to change something about the Range that the UDF is referring to or being called from. You think: "I'd just love to color this cell red if the UDF detects some state while executing". This thought comes up because it'd be nice to have the UDF update when called and even better if oyu can avoid dealing with conditional formatting. Alas, this is not allowed. The UDF must execute without making a change to the spreadsheet. This generally makes sense if you think about how Excel goes about calculating the spreadsheet. It makes a map of how cells are related and then proceeds to calculate the values in an order where each cell that depends on another is calculated in the precise order that is required. This process allows Excel to complete as fast as possible, without errors, and while using as many CPU cores as possible. If your UDF is able to change the spreadsheet after Excel has determine the order of calculations, then it becomes impossible to ensure that the spreadsheet is still correct. Because of this, Excel does not allow sde effects from a UDF.

The other aspect of this limitation that comes up often enough in practice ist hat you cannot use a Worksheet function that modifies the spreadsheet even if you intend to undo that function. For example: I have attempted to use the AutoFilter inside a UDF in order to determine how many times some condition showed up in a table. This is not allowed even though I intended to undo the AutoFilter before returning from my UDF. This limitation also applies to Copy/Paste and other common functions.