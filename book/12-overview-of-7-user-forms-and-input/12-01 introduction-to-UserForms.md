## introduction to UserForms

This chapter will focus on how to use UserForms to create interface that allow the user to interact with your VBA code.  UserForms can be used from anything to simple text inputs to very complicated forms.  There is really no limit to what you can do with UserForms, but at some point you will hit the limit of what you want to do inside the VBE.  Some folks will push the limit and develop fully featured programs in Excel.  I'd highly recommend you not do that and instead use UserForms to augment a good usage of VBA with useful interactivity.

When considering whether or not to use UserForms, there are a handful of pros and cons to using them.

Pros to using UserForms

* Provide a form to display/edit user input independent of Excel
* Provide for much better interaction with the user via "normal" programming events click, keyboard input, etc.
* Allow your program to collect several pieces of information before completing an action, especially if some real time process of the information is useful
* Provide a form that can "sit" on top of Excel and provide helper functionality or application specific functionality

Cons to using UserForms

There are a number of alternatives to creating a UserForm that are worth considering before committing to UserForms for a specific application.  Those alternatives and the cons against UserForms include:

* Editing the code of a UserForm can be a bit of a nuisance because you have to flip between design and code views
* The InputBox provides a simple way to collect a number of different input types without needing to create your own form.  For simple inputs (including Ranges), the InputBox is typically simpler and more consistent across applications
* Some types of inputs (including lists) can be easier to manage in a non-VBA context by using default Excel features.  For example, you do not need a ListBox on a UserForm if you can put a Table somewhere in the spreadsheet.
* If you are trying to use some form of version control on your source code, UserForms are very difficult to manage and version.
* If you want to provide buttons to perform an action, the Ribbon can be much more robust.  Of course editing the Ribbon can be a different pain, and I've gone the other way on this point before.

Once you've decided you want to use a UserForm to provide for user input/interaction, there are a number of areas that are important.  Those areas from start to finish are:

* Adding a UserForm to an existing Workbook or addin
* Using a Sub to make a UserForm show up using a keyboard shortcut or button or some other means
* Adding controls to UserForm and using those inputs to drive VBA
* Working with a UserForm to process input and update the UserForm
* Using the non-control events to provide interactivity
