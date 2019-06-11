### Processing User Input

User Input on a User Form is one of the most critical aspects ofm aking them. IT is less common to use a Form purely for output of information (although that is done). typically, you use a Form to provide input in a format that is easier to use than the default Excel interface. There are a ahdnful of Controls which are viewed as collectors of user input. You can then process than input in an Even tHandler or in other code which accesses the properties of the Control. Those common controls are:

- TextBox: Works great when you want to control a single value from the user. You can then parse the string into a number or whatever else you need
- ListBox.: Works great for allowing the user to select form a list from still being able to see multiple items in the list. Also supports multiple selection
- COmboBox: Same as ListBox but the control collapses to a single line when you are not selecting items. Does not allow for multiple selection.
- CheckBox or RadioButton: Allow the user to make a selection between choices while seeing the choices
- Button: Really allows a user to input a single click
- RefEdit: Not recommended but it allows you to select a Range from the Spreadsheet.
- TODO: any others (number bumper?)

For each of those Controls, you have a number of events which can be used to process the input as it comes in, or you can process the Properties of the Control once other code is running. One common pattern is to allow the user to input data into a number of TextBoxes, hit a button to run some action, and ten process all of hta input in one step after the button press. Another way to do the same thing would be to process and validate the input as it comes in, providing an error message if bad data was input.

For most of the Controls given above, you will find a `Value` property which gives either the Text of the Control or the selected state. The one exception to this is the ListBox which requires a little more work to get the Selection. For the ListBox, you need to iterate the items and check if the `Selected(index)` property of the ListBox is `True`.

TODO: add an example of using Value

Once you have the user input, it will typically be a `String` or a `Boolean`. To do something with these inputs, you will need to parse them into the desired types if not a string. The most common transformation is to parse a number from the string. This is done with `CInt` or `CDbl` which will *C*convert a String into a Integer or Double. You will get an error if the string was not parseable. If you do not need a number, there are a couple of other "C" functions:

- CBool
- CDate
- CErr
- TODO: add others, and descriptions
