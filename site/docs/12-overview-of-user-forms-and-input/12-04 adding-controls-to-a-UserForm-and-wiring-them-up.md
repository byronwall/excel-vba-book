## adding controls to a UserForm and wiring them up

Once your UserForm is created and the defaults are changed, your next task is to do something useful with the form. To that end, you will quickly need to add controls to the form and then wire those forms to useful actions. Adding the controls to the form is a straight forward process: show the Toolbox and then drag the items onto the form. Once you have dragged out a button, text box, and possibly a ListBox, you can simply copy and paste the previous items and avoid the Toolbox all around. There are a handful of controls that are not in the Toolbox by default. My strong advice here is to not use those controls if you are deploying this addin. Inevitably, some user will not have the OCX file or whatever is required to make it work. Just pass. Having said that, you may need to add a Date picker or possibly a RefEdit which are not included here (TODO: is that true about RefEdit?).

Dragging a control onto the form is fairly easy compared to the actual task of making a control do something useful. Below is a quick primer on how the different controls work, which properties are important, and will give you a guide for accomplishing 90% of what can be done with forms.

### CommandButton

The CommnadButton or simply button is one of the most common controls to use. Its use is simple, known to everyone, and easy enough to program against. A button does one thing: get clicked. The event you want to know about is the `_Clicked` event. Fortunately, the VBE will automatic create and wire up this event if you double click the button on the Designer version of the form. This makes it dead simple to create the button code that you want: just double click the button.

Note that the default event will be created with the current name of the control. To avoid this, you need to change the name of the button before you create the event. Be aware that VBA and the VBE are not that smart with respect to naming things nad wiring up changes. If you change the name of the button after you create the event, your event will not work. You should not change the button name after creating the vent (or plan to recreate it).

Other properties of the button that might be used:

- Value - will change the text that is displayed (TODO: is this right?)
- Enabled - will change whether the button can be pressed and will change the visuals. Useful when you want to show that an option could be possible but is not currently allowed or enabled
- Formatting and other visuals - you may change this on the property editor but it is far less common to modify the formatting once you are running. It can be done but is not common.

That's it; buttons are simple.

### CheckBox and Radio

The CheckBox and Radio are cousins (or siblings?) of each other and will be dealt with at once. They allow for a Boolean selection of an option. For the Checkbox, you are allowed to indicate the on/off state of a given button. For a Radio, you are allowed to indicate the on/off state for a single option _within a group of options_. The main thing to note about the Radio is that by selecting one item, you will deselect the others. In this way, the uses of these two controls maps naturally to the tasks you are likely to see.

Aside from the Name, the main items to deal with are:

- Clicked event - just double click to get this one
- Value - note you get this by default using the name, but it will include a Boolean of the selected state
- Enabled - can be used to disable the control

That's about it. You can change the formatting and other stuff, but these items typically exist to get an input and get to the real work. They are very common when you are providing options to the user or otherwise want to direct downstream If/Switch statements.

Beware that the Click event may be changed multiple times depending on how it was triggered (TODO: is that right?).

### TextBox

The TextBox is another simple one: it provides a means for the user to provide some text input. They work great for a range of things including input and output, although input is more typical. The idea is simple, the user provides a string and you use it somewhere. The properties to know:

- Value - this gets or sets the value that is displayed
- Enabled - can be used to disable the control (TODO: same as readonly?)

In terms of events, the main one to watch for is the KeyPress (TODO: or changed?). The idea is simple, if you want to track the input of the user, you tag along for that event and can respond to their key presses. The common uses of this are:

- Close a form or clear an input when ESC is pressed
- Do some action when ENTER is pressed
- Provide some form of validation or checking as the user types to either modify their input (e.g. ignore dashes) or otherwise update the UI based on their input.

TODO: add some addl onctent here about the event and its callback/parameters

That's it.

### ListBox

The ListBox is one control that has a number of options and a means of using it that are less obvious than the other controls. It's a shame really that the ListBox is so unintuitive in VBA because it is quite powerful and other programming languages have handled this better. THe idea behind a ListBox is that it provides a list of items whose use can vary according to what you want. Some common applications include:

- Allow the user to select from one or multiple options in a list
- Provide some output to the user (and possibly then use that output as the input for next step)

THe input/output decision here is somewhat critical because the things that will annoy you about the ListBox break on this point. If you are collecting input, then really you have to also deal with output because at the end of the day, you have to put something in the ListBox in order for a user to select it. Once you've handled the output stuff, then determining which items have been selected by the user is straightforward enough. Therefore, covering the output part is a good starting point.

To put items into the ListBox, you need to modify the List collection on the object. There are two ways to do this:

- Directly, via the List object
- Indirectly, using the `AddItem` command

Either way you go, you have a couple of decisions after adding the item: what text do you want displayed for the item and do you want multiple columns? If you are dealing with a single column, then you can simply add the text in the call for an addition and that's all. IF you are working with columns, then you will need to do two things:

- Set up the columns (using the editor or via commands) (TODO: add pictures or code here)
- Call the command to set the fields using the row and column number (TODO: add some code)

ALthough I have described a simple process here, oftentimes, you will deal with something that is more complicated. THe issue comes when you want to maintain some reference to an object but you are required to use a string for display purposes. This means that you need some means of maintaining that reference back to the object. There are options for dealing with this:

- Rely on the index of the objects matching (and not changing) and simply use the row index
- Create a Dictionary that stores the link between the string and the object
- Use some other object or Collection that can reference the object back to the string
- Serialize the object into the ListBox value (if multiple fields, join with a `|` or similar)

Each of those approaches has its pros and cons, btu the main idea is that you are often forced to deal with something that is typically much easier in other languages. My general approach is to rely on row index if I know that changes are not possible. This is common for a lot of code since yout ar likely to control both side. If that is not ideal, then you can typically find some way to store a reference between the display value and the object using a Dictionary.

Once you hav eth information in the ListBox, you can simply iterate the `Items` by index nad check the `Selected(index)` property to see if the item is selected. Note that if you do not allow multiple selection, then you can also use the `SelectediNdex` property (TODO: is that right?).

TODO: add some code here to demonstrate iterate through a ListBox

Although this section has the most text, the ListBox is not always a pain to deal with. Typically they are much better than the alternatives (like using the Excel spreadsheet somehow) but require that you remember some boilerplate for accessing and changing items.

### Other Controls

There are a couple of other controls that you may see that are summarized here:

- Label: these don't do much other than provide some fixed text when could be changed later (I rarely ever do it)
- RefEdit: this control technically allows you to select a Range from Excel. They are quite buggy. Depending on you main goal, you may of much better to use `Application.InputBox(Type:=8)` to access a Range.
- Tabs: these can be helpful for organizing a complicated workflow. You will find yourself wanting to change the active tab and possibly limit access to later tabs.
- Wells?, whatever it's called, there allow you to Group controls. These may be required for a Radio to work like you want (if you have multiple sets of Radios on a single form).
