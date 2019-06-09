### Processing User Input

User Input on a User Form is one of the most critical aspects ofm aking them. IT is less common to use a Form purely for output of information (although that is done). tpyically, you use a Form to provide input in a format that is easier to use than the default Excel itnerafce. There are a ahdnful of Controls which are viewed as collectors of user input. You can then process than tinput in an Evne tHandler or in other code which accesses the properties of the Control. Those common controls are:

- TextBox: Works great when you want to control a single value from the user. You can then parse the string into a number or whatever else oyu need
- ListBox.: Works great for allowing the user to select form a list from still beign able to see multiple items in the list. Also uspports multiple seleciton
- COmboBox: Same as ListBox but the control collapses to a single line when you are not selecting items. Does not allow for multiple selection.
- CheckBox or RadioButton: Allow the user to make a selection between choices while seeing the choiecs
- Button: Really allows a user to unput a single click
- RefEdit: Not recommended but it allows you to select a Range from the Spreadsheet.
- TODO: any others (number bumper?)

For each of those Controls, you have a number of evnets which can be used to process the input as it comes in, or you can process the Properties of the Control once other code is running. One common pattern is to allow the user to input data into a number of TextBoxes, hit a button to run some action, and ten process all of hta tinput in one step after the button press. Anotehr way to do the same thing would be to process and vlaidate the input as it comes in, providign an error message if bad data was input.

For most of the Controls given above, you will find a `Value` property which gives either the Text of the Control or the selected state. The one exception to this is the ListBox which requires a little more work to get the Selection. For the ListBox, you need to ierate the items and check if the `Selected(index)` property of the ListBox is `True`.

TODO: add an example of using Value

Once you have the user input, it will typically be a `String` or a `Boolean`. To do something with these niputs, you will need to parse them into the deisred types if not a string. The most ocmmon transformation is to parse a number from the string. This is done with `CInt` or `CDbl` which will *C*onvert a String into a Integer or Double. You will get an error if the string was not parseable. If you do not need a number, there are a couple of other "C" functions:

- CBool
- CDate
- CErr
- TODO: add others, and descriptions
