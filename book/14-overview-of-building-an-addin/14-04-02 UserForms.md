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
