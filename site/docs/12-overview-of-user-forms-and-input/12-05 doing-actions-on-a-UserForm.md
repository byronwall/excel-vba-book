## doing actions on a UserForm

Section will focus on performing commands on a UserForm without leaving the UserForm. This looks mostly like normal forms programming.

Programming on a UserForm is one area of VBA that is largely independent of Excel. Yes, you are able to access the Object Model from a form, but most of the programming on a form is simple related to the form. This is also the one area of VBA where if you have done it before in another language, your experience will transfer nearly 1:1. Forms programming is largely all about building a good UI that meets your needs. You can do a lot with VBA in terms of events and other sorts of dynamic programming, but most of the itme you just make the form and get on with the real work.

Some of the comment things you will want to of when programming on a form include:

- Creating event handlers
- Parsing and responding to user input
- Populating data in the UserForm from the Object Model
- Accessing Properties of controls to change or use

### Event Handlers

Event Handlers are at the core of User Forms and making them useful. To be clear, your Form will do nothing without events. You could it to display static content from the designer mode, but it will do nothing useful. To make your Form become useful, you add Controls to it and then add Events to those Controls. Event Handlers are the glue (or wires) that take the actions performed on Controls and direct them somewhere useful. Events control everything from Clicking, Loading, Typing and everything else. Each Control has a unique set of events depending on what it can do, but in general, there's a bit of overlap between different controls.

To add an event handler, there are a couple of options:

- Double click on the Control in Design Mode, and you will get the default event handler created
- Go to the code view, and select the Control and then Event you want from the drop downs (TODO: add image)
- Type the Event handler based on the named of the Control and the event you want

If you know the default events, then option 1 is as good as the theories. IF you want to see a list of events before creating one, then you will go with option 2. You will pretty much never type the event handler out by handler unless you are copying it from somewhere else.

Once you have created the handler, you simply add the code that you want to fire in the event. One good tip here is to use the event handler to call other Subs. It's a good habit to not put logic or other execution based code into Even thNadlers. The reason for this is that you may want to perform the same action from multiple events. Putting the code in a handler makes it difficult to reuse the code because some ehadnlers have parameters nad other details that make it hard to arbitrary call them. Of course, I regularly put code into event Handlers, but at least I know I should avoid it. I am constantly reminded of why to avoid it when I have to extract code from one event to put into a Sub to call from another event.

One important note about Event handlers is that the handler can have some number of parameters that are included in the handler signature. These parameters ar etpyically used to pass along information related to the event. For example the key press event contains the key code of the key that was pressed. The Click event however has no parameters. The presence of parameters is easy to check when the VBE creates the handler for an event since it will give the parameters.

TODO: given an example of using Handlers?

TODO: include a blurb about the Initialize event (if it was not addressed earlier)

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

### Accessing the Excel Object Model

From a UserForm, you have full access to the Excel Object Model. This can be very handy if you are trying to access information from the USerForm to determine what information to show in the Form. It can also be helpful if you want to make changes to the underlying spreadsheet from a USerForm without leaving the form. Both of those options are very common and very easy to do with UserForms. In general, any code that can run without a USerForm present can be run with a USerForm. There are some limitations when it comes to the user's ability to Select items with a From visible, but you are not limited in calling the same commands from VBA (TODO: is that right?). The exception ehre is that if the form is `ShowModal = False` then the user is able to make selections while the from is visible.

There is no real limit to what you can do from a SuerForm. A couple of examples to give you a feel:

- present a list of all open Workbooks so that they user can select which one that want to process
- Create a form that can process all of the selected Charts.
- Present a ListBox with the unique values from all of the AutoFilters that are active. Allow the user to selectively remove or change those filters without having to use the normal drop downs.

### Accessing control Properties

THe final piece of Forms programming is somewhat meta: allow the UserFOrm code to change the USerFOrm. There are a couple of obvious reasons you might want to do this:

- Change the position of the USerFOrm (center on start)
- Enable or disable a button or other control based owns ome input. You can extend this to making things visible or not as well.
- Change the text, format, or other visual detail of a Control based on some other state or user input.

TODO: add the code for centering a UserForm.

IN addition tot hose simple conners, you also have the ability to danmically create controls on demand. This makes it possible to add/remove controls to the USerForm as needed. This can be helpful if you want to create Control based on some proeprt yof the Worksheet but where you may not know how many times to do it in advance. For example. maybe you want to provide a LIstBOx with unique values for each column that was selected. IN advance, you may not know the column count so you need to create ListBoxes on demand. This can be done with UserForm programming.

TODO: example of create a Control from scratch
