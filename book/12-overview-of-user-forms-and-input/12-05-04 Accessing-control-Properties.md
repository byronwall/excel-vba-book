### Accessing control Properties

THe final piece of Forms programming is somewhat meta: allow the UserFOrm code to change the USerFOrm. There are a couple of obvious reasons you might want to do this:

- Change the position of the USerFOrm (center on start)
- Enable or disable a button or other control based owns ome input. You can extend this to making things visible or not as well.
- Change the text, format, or other visual detail of a Control based on some other state or user input.

TODO: add the code for centering a UserForm.

IN addition tot hose simple conners, you also have the ability to danmically create controls on demand. This makes it possible to add/remove controls to the USerForm as needed. This can be helpful if you want to create Control based on some proeprt yof the Worksheet but where you may not know how many times to do it in advance. For example. maybe you want to provide a LIstBOx with unique values for each column that was selected. IN advance, you may not know the column count so you need to create ListBoxes on demand. This can be done with UserForm programming.

TODO: example of create a Control from scratch
