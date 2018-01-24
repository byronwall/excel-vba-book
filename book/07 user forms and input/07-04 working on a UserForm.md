## doing actions on a UserForm

Section will focus on performing commands on a UserForm without leaving the UserForm.  This looks mostly like normal forms programming.

Programming on a UserForm is one area of VBA that is largely indepednet of Excel.  Yes, you are able ot access the Object Model from a form, but most of the programming on a form is simple related to the form. This is also the one area of VBA where if you have done it before in another language, your expereince will transfer nearly 1:1.  Forms programming is largely all about building a good UI that meets your needs.  You can do a lot with VBA in terms of events and other sorts of dyanmic programming, but most of the itme you just make the form and get on with the real work.

Some of the ocmmont things you will want ot od when programming on a form include:

* Creating event handlers
* Parsing and respoding to user input
* Populating data in teh UserForm from the Object Model
* Accessing Properties of controls to change or use

### Event Handlers

Event Handlers are at the core of User Forms and making them useful. To be clear, your Form will do  nothing without events.  You could it to display static content from the designer mode, but it will do nothing useful.  To make your Form become useful, you add Controls to it and then add Events to those Controls.  Event Handlers are the glue (or wires) that take the actions perofmred on Controls adn direct them somewhere useful.  Evnets control everything from Clicking, Loading, Typing and everything else.  Each COntrol has a unique set ofo events depending on what it can do, but in general, there's a bit of overalp between different controls.

To add an event handler, there are a couple of options:

* Double click on the Control in Design Mode, and oyu will get the defualt event handler created
* Go to the code view, and select the Control and then Event you want from the drop downs (TODO: add image)
* Type the Event handler based on teh named of the Control adn the event you want

If you know the defualt events, then option 1 is as good as teh toehrs.  IF you want to see a list of events beffore creating one, then you will go with optojn 2.  You will pretty much never type the event handler out by handler unless you are copying it from somewhere else.

ONce you have created the ahndler, you simply add hte code that oyu want to fire in teh event.  One good tip here is to use the event handler to call other Subs.  It's a good habit to not put logic or other execution based code into Even thNadlers.  The reason for this is that you may want to perfomr the same action from multipl events.  Putting the code in a handler makes it idfficult to resue the code becasue som ehadnlers have parameters nad other details that make it hard to arbitrary call them.  Of course, I regualrly put code into event Hnadlers, but at least I know I shoudl avoid it.  I am constantly reminded of why to avoid it when I ahve to extract code from one event to put into a SUb to call from another event.

One important note about Event handlers is that the hadnler can have some number of parmaeters that are included in teh handler signature.  These parameters ar etpyically used to pass along infomraiton related to the event. For ecample the key press event contains the key code of the key that was pressed.  The Click event however has no paramteers.  The presence of apramters is easy to check when teh VBE creates the handler for an event since it will give the parameters.

TODO: given an example of using Handlers?

TODO: include a blurb about the Initialize event (if it was not addressed ealrier)

### Processing User Input

User Input on a User Form is one of the most critical aspects ofm aking them. IT is less common to use a Form purely for output of information (although that is done).  tpyically, you use a Form to provie input in a format that is easier to use than teh default Excel itnerafce.  There are a ahdnful of Controls which are viewed as collectors of user input.  You can then process tha tinput in an Evne tHandler or in other code which accesses hte properties of the Control.  Those common controls are:

* TextBox: Works great when you wnat to control a single value from teh user.  You can then parse the string into a number or wahtever else oyu need
* ListBox.: Works great for allowing the user to select form a list from still beign able to see multiple items in teh list.  Also uspports multple seleciton
* COmboBox: Same as ListBox but the contorl collapses to a single line when you are not selecting items.  Does not allow for multiple selection.
* CheckBox or RadioButton: Allow the user to make a sleection between choices while seeing the choiecs
* Button: Really allows a user to unput a single click
* RefEdit: Not recommended but it allows you to select a Range from the Spreadsheet.
* TODO: any others (number bumper?)

For each of those Controls, you have a number of evnets which can be used to process the input as it comes in, or you can process the Properties of the Control once other code is running.  One common pattern is to allow the user to input data into a number of TextBoxes, hit a button to run some action, and ten process all of hta tinput in one step after hte button press.  Anotehr way to do the same thing would be to process and vlaidate the input as it comes in, providign an error message if bad data was input.

For most of the Controls given above, you will find a `Value` property which gives either the Text of the Contorl or the selected state.  The one exception to this is the ListBox which requries a little more owrk to get the Selection.  For the ListBox, you need to ierate the items adn check if hte `Selected(index)` property of the ListBox is `True`.

TODO: add an example of usign Vaule

Once you have the user input, it will typically be a `String` or a `Boolean`.  To do somethign with these niputs, you will need to parse them into the deisred types if not a string.  The most ocmmon transformation is to parse a number from teh string. This is done with `CInt` or `CDbl` which will *C*onvert a String into a Integer or Double.  You will get an erorr if the string was not parseable.  If you do not need a number, there are a couple of other "C" functions:

* CBool
* CDate
* CErr
* TODO: add others, and descriptions

### Accessing the Excel Object Model

From a UserForm, you have full access to the Excel Object Model. Thsi can be very handy if you are trying to access informaiton from teh USerForm to determine what infomraiton to show in teh Form.  It can also be helpful if oyu want to make changes to teh underlying spreadsheet from a USerForm without leaving the form.  Both of those options are very common and very easy to do with UserForms.  In general, any code that can run without a USerForm present can be run with a USerForm.  There are some limiations when it comes to teh user's ability ot Select items with a Fomr visible, but you are not limited in calling teh same commands from VBA (TODO: is that right?).  Teh exception ehre is that if the form is `ShowModal = False` then the user is able to make selections while the fomr is bisible.

There is no real limit to waht you can do from a SuerForm.  A couple of examples to give you a feel:

* present a list of all open Workbooks so that they user can sleect which one that want to process
* Create a form that can process all of the slected CHarts.
* Present a ListBox with teh unique values from all of the AutoFilters that are active.  Allow the user to selectively remove or chagne those filters without having to use the normal drop downs.

### Accessing control Properties

THe final piece of Forms programmign is somewhat meta: allow the UserFOrm code ot change the USerFOrm.  There are a couple of obvious reaosns you might want to do this:

* Change the position of the USerFOrm (center on start)
* Enable or disable a buttn or other control based ons ome input.  YOu can extend this to making things vivsible or not as well.
* Change teh text, format, or other visual detail of a Control based on some other state or user input.

TODO: add the code for cnetering a UserForm.

IN addition tot hose simple concners, you also have hte ability ot danmically create controsl on demand.  This makes it possible to add/remove controls ot the USerForm as needed.  This can be helpful if oyu want to create Control based on some proeprt yof the Worksheet but where you may not know how many times to do it in advance.  For example. maybe you want to prvoide a LIstBOx with unque values for each column that was slected.  IN advnace, you may not know the column coujtn so you need to create ListBoxes on demand.  This can be done with UserForm programming.

TODO: example fo create a Control from scrathc
