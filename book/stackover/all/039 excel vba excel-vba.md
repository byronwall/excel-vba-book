# SO item 039
I am trying to make a red label appear on my form and can not. I have tried changing the property to RGB and HEX and just get errors. Is there a way to get a property value to make my label RGB(200, 0, 0)? I am unaware of how the value in the property areas is developed.

This is the only way I can make a red label:

```
Private Sub Label13_Click()
 Label13.BackColor = RGB(200, 0, 0)
End Sub

```

I have to click the label to make it red. Is there a way to use code to make it red when the form starts? Or perhaps generate a value for the property? Thank you for your help in advance.

----

You can use the `Initialize` event for the form instead of putting the event on a `Click` event.

Here is an example with the form named `UserForm`. Use the dropdowns to select the form and then the `Initialize` event.

![form events](https://i.stack.imgur.com/NY9fM.png)

You can also just set the color in the properties if you know this is the color you want.

![label properties](https://i.stack.imgur.com/JL8sr.png)
