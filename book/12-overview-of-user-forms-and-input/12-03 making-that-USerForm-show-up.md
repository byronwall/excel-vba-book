## making that USerForm show up

Once your UserForm is created, there are a couple of ways of showing it on screen:

- Run any code from the VBE that is contained within the form. This will show the form.
- Create an instance of the form somewhere and show it

For those two methods, the latter is really the only one that will work for user applications or other "real" uses. If you are simply testing or doing things for yourself, then hitting F5 in the VBE may not be a large ask.

For the former, see the code below for an example of how to show the form.

```vb
DIm frm as UserForm
Set frm = New UserForm

frm.Show
```
