## adding controls to a UserForm and wiring them up

Once your UserForm is created and the defaults are changed, your next task is to do somethign useful with the form.  To that end, you will quickly need to add controls to the form and then wire thsoe forms to useful actions.  Adding the contorls to the form is a straight forward process: show the Toolbox and then drag the items onto the form.  Once you have dragged out a button, text box, and possibly a ListBox, you can simply copy and paste the previous items adn avoid the Toolbox all around.  There are a handful of controls that are not in the Toolbox by default.  My strong advice here is to not use those controls if you are deploying this addin.  INeveitably, some user will not have the OCX file or whatever is requried to make it work.  Just pass.  Having said that, you may need to add a Date picker or possibly a RefEdit which are not included here (TODO: is that true about RefEdit?).

Dragging a control onto the form is fairly easy compared to the actual task of making a control do somethign useful.  Below is a quick primer on how the idffernet contorls work, which properties are important, and will give you a guide for accomplishing 90% of waht can be done with forms.

### CommandButton

The CommnadButton or simply button is one of the most common controls to use.  Its use is simple, known to everyone, and easy enough to program against.  A button does one thing: get clicked. The event you want to know about is the `_Clicked` event.  Fortunately, the VBE will atuomtallica create and wire up this event if you double click the button on the Desinger versiojn of the form.  This makes it dead simple to create the button code that you want: just double click the button.

Note that the default event will be created with the current name of the control.  To avoid this, you need to change the name of the button before you create the event.  Be aware that VBA and the VBE are not that smart with respect to naming things nad wiring up changes.  If you change the name of the button after you create the event, your event will not work.  You should not chagne teh button name after creating the vent (or plan to recreate it).

Other properties of the button that might be used:

* Value - will change the text that is dispalyed (TODO: is this right?)
* Enabled - will change wheteher the button can be pressed and will change the visuals.  Useful when you want to show that an option could be possible but is not currently allowed or enabeld
* Formatting and other visuals - you may change this on the proeprty editor but it is far less common to modify the fomratting once you are running.  It can be done but is not common.

That's it; buttons are simple.

### CheckBox and Radio

The CheckBox and Radio are cousins (or siblings?) of each other and will be dealt with at once.  They allow for a Boolean selection of an option. For the Checkbox, you are allowed to indicate the on/off state of a given button.  For a Radio, you are allowed to indicate the on/off state for a singel option *within a group of options*.  The main thing to note about the Radio is that by selecting one item, you will uneslect the others.  In this way, the uses of these two controls maps naturally to the tasks you are likely to see.

Aside from the Name, the main items to deal with are:

* Clicked event - just double click ot get htis one
* Value - note you get this by defualt usign the name, but it will include a Boolean of the selected state
* Enabled - can be used to disable the control

That's about it.  You can change the formatting and other stuff, but these items typically exist to get an input and get to the real work.  They are very common when you are providing options to the user or otherwise want to direct downstream If/Switch statements.

Beware that the Click event may be changed multiple times depending on how it was triggered (TODO: is that right?).

### TextBox

The TextBox is another simple one: it provides a means for the user to provide some text input.  They work great for a range of things including input and output, although input is more typical.  The idea is simple, the user provides a string and you use it somewhere.  The properties to know:

* Value - this gets or sets teh value that is displayed
* Enabled - can be used to disable the control (TODO: same as readonly?)

In terms of events, the main one to watch for is the KeyPress (TODO: or changed?).  The idea is simple, if oyu want to track the input of hte user, you tag along for that event and can respond to their key presses.  The common uses of htis are:

* Close a form or clear an input when ESC is pressed
* Do some action when ENTER is pressed
* Provide some form of vlaidation or checking as the user types to either modify their input (e.g. ignore dashes) or otherwise update the UI based on tehir input.

TODO: add some addl onctent here baout the event and its callback/parameters

That's it.

### ListBox

The ListBox is one control that has a number of options and a means of using it that are less obviosu than the other controls.  It's a shame really that the ListBox is so unintuitive in VBA beucase it is qutie powerful and other programming langugaegs have handled htis better.  THe idea behidn a ListBox is that it provides a list of items whose use can vary according to what you wnat.  Some common applications include:

* Allow the user to select from one or multiple options in a list
* Provide some output to the user (and possinly then use that output as the input for next step)

THe input/output decision here is somewhat critical because the thigns that will annoy you about the ListBox break on this point.  If you are collecting input, then really you have to also deal with output because at the end of the day, you ahve to put somethign in teh ListBox in order for a user to select it.  ONce you've handled the output stuff, then determinign whcih items have been sleected by the user is straightforward enough.  Therefore, covering the output part is a good starting point.

To put items into teh ListBox, you need to modify the List collection on the object.  There are two ways to do this:

* Directly, via the List boject
* Indirectly, usign the `AddItem` command

Either way you go, you ahve a couple of decisions after adding the item: what text do you want displayed for the item and do you want multiple columns?  If you are dealing with a singel column, then you can simply add the text in the call for an addition and that's all.  IF you are working with columns, then you will need to do two things:

* Set up the columns (using the editor or via commands) (TODO: add pictures or code here)
* Call the commnad to set the fields usign the row and column nujmber (TODO: add some code)

ALthough I have descibed a simple process here, oftentimes, you will deal with something that is more complicated.  THe issue comes when you want to maintain some reference to an obiject but you are reuqired to use a string for display purposes.  This means that you need some menas of maintaining that refernece back to the object.  There are options for dealing with this:

* Rely on teh index of the objects matching (and not changing) and simply use the row index
* Create a Dictionary that stores the link between the string and the object
* Use some other object or Colleciton that can reference the object back to teh string
* Serialize teh object into the ListBox value (if multiple fields, join with a `|` or similar)

Each of those approaches has its pros and cons, btu the main idea is that you are often forced to deal with somethign that is typically much easier in other languages.  My general appproach is to rely on row index if I know that changes are not possible.  This is common for a lot of code isnce yout ar elikely to contorl both side.  If that is not ideal, then you can typucally find some way to store a referenc ebetween the display value and the object usign a Dictionary.

Once you hav eth einfomration in the ListBox, you can simply iterate the `Items` by index nad chekc the `Selected(index)` property to see if the item is sleected.  Note that if you do not allow multiple selection, then you can also use the `SelectediNdex` property (TODO: is that right?).

TODO: add some code here to demonstrate iterate through a ListBox

Although this section has the most text, the ListBox is not always a pain to deal with.  Typically they are much better than the alternatives (like using the Excel spreadsheet somehow) but require that you remember some boilerplate for accessing and changing items.

### Other Controls

There are a couple of other controls that you may see that are summarized here:

* Label: these don't do much other htan provide some fixed text when could be changed later (I rarely ever do it)
* RefEdit: this contorl tehcnically allows you to select a Range from Excel.  They are quite buggy.  Depending on you main goal, you may od much better to use `Application.InputBox(Type:=8)` to access a Range.
* Tabs: these can be heplful for organizing a complicated workflow.  You will find yourself wanting to change the active tab and possibly limit access to later tbas.
* Wells?, whatever it's called, there allow you to Grouo controls. These may be requried for a Radio to work like you want (if you have mutliple sets of Radios on a singel form).
* 