## creating a UserForm

Creating a UserForm is a simple process: open the VBE and then right click to Insert -> UserForm.  This will give you a default UserForm that is blank, has a default name, and is a default size.  If you're lucky, this will also open the UserForm for editing and show you the toolbox which provides controls to edit.  Once the form is created, there are a couple of things which you should do immediately, before you forget:

* Change the name of the form to something more useful than UserForm1
* Change the `Caption` on the form to something better than UserForm
* Consider changing the `ShowModal` property if you know you do not want a modal dialog

Once those items are done (or decided against) you can start adding controls to the USerForm and changing its size.

TODO: add pictures of the steps

## making that USerForm show up

Once your UserForm is created, there are a couple of ways of showing it on screen:

* Run any code from the VBE that is contained within the form. This will show the form.
* Create an instance of the form somewhere and show it

For those two methods, the latter is really the only one that will work for user applicatiojns or other "real" uses.  If you are simply testing or doing things for yourself, then hitting F5 in the VBE may not be a large ask.

For the former, see the code below for an example of how to show the form.

```vb
DIm frm as UserForm
Set frm = New UserForm

frm.Show
```
