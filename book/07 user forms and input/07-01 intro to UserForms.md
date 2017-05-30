## introduction to UserForms

This chapter will focus on how to use UserForms to create interface that allow the user to interact with your VBA code.  UserForms can be used from anything to simple text inputs to very complicated forms.  There is really no limit to what you can do with UserForms, but at some point you will hit the lmit fo waht you want to do inside the VBE.  Some folks will push the limit and develop fully featured programs in Excel.  I'd hihgly recommend you not do that and instead use UserForms to augment a good usage of VBA with useful interactivity.

When cosnidering whether or not to use UserForms, there are a handful of pros and cons to using them.

Pros to using UserForms

* Provide a form to display/edit user input independent of Excel
* Provide for much better interaction with the user via "normal" programming events click, keyboard input, etc.
* Allow your program to collect several pieces of information before completing an action, espcially if some real time process of the information is useful
* Provide a form that can "sit" on top of Excel and provide helper functionality or application specific functionality

Cons to using Userforms

There are a number of alternatives to creating a UserForm that are worht considering before comitting to UserForms for a specific application.  Those alternatives and teh cons against UserForms include:

* Editing the code of a UserForm can be a bit of a nuisance because you have to flip between design adn code views
* The InpuxBox provdies a simple way to collect a number of different input types without needing to create your own form.  For simple inputs (inlucding Ranges), the InpuxBox is typically simpler adn more consistent across applications
* Some types of inputs (including lists) can be easier to manage in a non-VBA context by using default Excel features.  For example, you do not need a ListBox on a UserForm if you can put a Table somewhere in teh spreadsheet.
* If you are trying to use soem form of version control on your source code, UserForms are very difficult to manage and version.
* If you wnat to provide buttons to perform an aciton, the Ribbon can be much more robust.  Of course editing the Ribbon can be a different pain, and I've gone the other way on this point before.

Once you've decided you want to use a UserForm to provide for user input/interaciton, there are a number of areas that are imporatnt.  Thsoe areas from start to finish are:

* Adding a UserForm to an exisitng Workbook or addin
* Using a Sub to make a UserForm show up using a keyboard shortcut or button or some other means
* Adding controls to UserForm and using those inputs to drive VBA
* Working with a UserForm to process input and update the UserForm
* Using the non-control events to provide interactivity

