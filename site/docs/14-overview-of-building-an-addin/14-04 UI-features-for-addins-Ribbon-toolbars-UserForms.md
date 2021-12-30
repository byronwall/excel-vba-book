## UI features for addins, Ribbon, toolbars, UserForms

There are a number of UIs that can be provided for an addin. The most common involve using the Ribbon or providing UserForms (typically accessible from a keyboard shortcut). Those two approaches will discussed in detail. It is also worth mentioning that if you are going to support Excel 2007 and before, that the Ribbon did not exist back then. For those prior versions of Excel, the interfaces were built using toolbars and menu items. That's before my day, so I'll just say that if you use that code today, it will show up in the Ribbon. If you need to support those versions of Excel, you would do well to find a different book.

### the Ribbon

When using the Ribbon, there are a couple of items to consider:

- Do you want your UI to show up in an existing tab or on your own?
- How interactive do you want your UI to be? This can range from simple buttons that trigger actions to text boxes and other more interactive features that are able to detect user input and respond accordingly.
- How do you prefer to edit the file? How fast do you want your developments cycle to be with respect to the UI?

For the first point, this is a simple preferences. For a given addin it may make sense to simply put the buttons and other access on an existing tab (Developer and Data are popular!) and provide that level of access. For an addin that has a dedicated purposed independent of other Excel features, it starts to make sense to add your own tab exclusively for your addin. This is good for helping your users find your features. It can also be more consistent in terms of keyboard shortcuts. If you are going to modify an existing tab, be abolustley certain that you verify that the keyboard shortcuts work as expected. There is nothing worse than having an addin break the ALT+A+R+Y shortcut which is supposed to reapply an autofilter. It is not fun when that shortcut becomes ALT+A+R+Y2. Seriously?

For the second point, you will need to consider the average user and their expectations. Keep in mind that the default Excel Ribbon includes a number of locations where user input is collected nad used beyond a simple button. This includes things like some number input (font size, page layout) and other drop downs. There is a willingness for Excel users to ues these features where it make sense. For what it's worth, from the Excel VBA point of view, it is much simpler to not try nad collect user input. This can be done (TODO: add examples), but the effort here is typically not worth the user experience. If you choose to go this route, I would highly recommend using drop downs and other inputs that provide some automatic filtering of user input. Trying to validate user input off the Ribbon is a pain and does not provide a good experience. Having said that, if you are designing for power users, you can build a very slick interface in the Ribbon that is unmatched.

The final point gets down to the nitty gritty of actually editing the Ribbon. The problem is that the Ribbon is defined in a file inside your XLAM file and is not editable from any part of the VBE or Exce interface. This means that it is a real pain to edit the Ribbon definition in the same way that you can edit the other VBA code. I have typically taken the approach of using a button on the Ribbon to launch a form that exists in VBA. That form can then be edited without having to touch the Ribbon definiont. This sounds tribial, but it can make a huge difference if you are designing an addin that has a lot of possible interactivity; it is very difficult to edit the Ribbon in real time. Having said that, there is one addin that makes this process much more manageable. It is from Andy Pope (TODO: add link) and works great for building out the interface. Even using that addin, it is a pian to add the callback necessary to tie the Ribbon to VBA. Don't be dissuaded from creating a nice Ribbon UI, but realize that it takes time and effort and attention to detail to properly detail out the Ribbon UI aspects.

#### editing the Ribbon

Note that the Ribbon is defined in an XML file inside your XLAM file. Remember that an XLAM file is simply a ZIP file of a bunch of different folders. By default, the Ribbon definition is not included and you must add the folder and file. To do this, simple create the `customUi` folder and then create a XXX XML file inside there. This file will define the specific changes you are making to the Ribbon.

#### callbacks

Once you hav the Ribbon XML set up, you will be defining callbacks that need to actually exist in the VBA code. I always like to create a `Ribbon` module in the XLAM file which is solely responsible for callbacks. This is nice for larger addins with a large number of callbacks because it provides a single place. It also avoids debugging errors later when you accidentally put some critical code into a callback and forgot to check that out.

The callbacks take an odd signature. I always use Andy Pope's addin or copy a previous one. I have very seldom used the parameters in the callback for accessing the Ribbon information. My approach has always been to avoid extra interactivity with the ribbon. I have done it before, and it works, but the problem is that it is just not intuitive to do it that way. It is much easier to add a keyboard shortcut which shows a USerForm than to attempt to get the user to focus on the Ribbon (using the keyboard or mouse) and then provide the best info.

### UserForms

If you are going to use UI features within your addin, you are going to use UserForms. They provide the cleanest and easiest interface for the vast majority of automation and other tasks. THe one exception to using a UserForm is when you can get by with a simple `InputBox`. You should alwa sprefer the INputBox because the procedure for calling them and obtaining a value is dead simple. Also, the InputBox is the best way to ask the user for a Range input via selection. You can technically use the `RefEdit` control, but hta control is very sensitive when it works.

If you are building a USerFOrm, there is very little than tis different from a normal USerForm. The only thing to be aware of are the logistics of creating, showing, and hiding the USeForm. I have previously tried to keep an instance of a given form live in order to use the previous values. This has worked well inside a single Workbook btu seems to be very finicky when working across multiple Workbooks. The procedure here is very simple:

```vb
DIm frm as UserForm
Set frm = New UserForm

fmr.Show
```

The code above is all that is required to create a new instance of a form and show it to the user. From there, the code is the same as before: you simply create the form and call the various Subs you want. One thing which is helpful is to hide the form when you are done. This is done with the `Unload Me` command.

One other item to be aware of is that the default UserForm is set to `ShowModal = True` which applies the modal property. A "modal" dialog is one who steals focus from any other elements and must be dealt with before you can go back to your previously focusbale elements. This is often good for certain workflows where you do not want the user to change the underlying spreadsheet while you collect their input. There are other instances however where it makes sense to allow the user to change the active Workbook, Worksheet, or Selection and then interact with you rform. To allow for this behavior, set `ShowModal = False`. THs will allow your form to exit even when the user clicks off and interacts with the spreadsheet again. This is a real game changer when you are working with code that operates on the current selection. You are then able to leave your form up while the user changes the selection. From there, they are able to call the cod ethey want on the objects they want. I have used this technique to great effect when working with Charts: allow the user to select their charts and then hit a button.
