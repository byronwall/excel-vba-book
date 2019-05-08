### TextBox

The TextBox is another simple one: it provides a means for the user to provide some text input. They work great for a range of things including input and output, although input is more typical. The idea is simple, the user provides a string and you use it somewhere. The properties to know:

- Value - this gets or sets teh value that is displayed
- Enabled - can be used to disable the control (TODO: same as readonly?)

In terms of events, the main one to watch for is the KeyPress (TODO: or changed?). The idea is simple, if oyu want to track the input of hte user, you tag along for that event and can respond to their key presses. The common uses of htis are:

- Close a form or clear an input when ESC is pressed
- Do some action when ENTER is pressed
- Provide some form of vlaidation or checking as the user types to either modify their input (e.g. ignore dashes) or otherwise update the UI based on tehir input.

TODO: add some addl onctent here baout the event and its callback/parameters

That's it.
