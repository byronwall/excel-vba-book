# SO item 028
I created a UserForm with several functions.

The form opens as I open the Excel file, however if I try to close the form the Excel file close together. Additionally, I can't open the VBA of this Excel (containing the form), so what I do (and that is really dum) it is to open another Excel, press <kbd>Alt</kbd>+<kbd>F11</kbd> to open the macro environment and then I can open my Excel file with the UserForm.

I think my problem is in this specific code:

```
Private Sub UserForm_Terminate()
    'Application.Visible = True
    ActiveWorkbook.Saved = True
    Application.Quit
End Sub

```

Can anyone guess what is the problem here?

----

If you just want to close the Userform use `Unload Me` instead of `Application.Quit`.
