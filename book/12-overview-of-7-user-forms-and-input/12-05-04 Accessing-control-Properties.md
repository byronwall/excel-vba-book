### Accessing control Properties

THe final piece of Forms programmign is somewhat meta: allow the UserFOrm code ot change the USerFOrm. There are a couple of obvious reaosns you might want to do this:

- Change the position of the USerFOrm (center on start)
- Enable or disable a buttn or other control based ons ome input. YOu can extend this to making things vivsible or not as well.
- Change teh text, format, or other visual detail of a Control based on some other state or user input.

TODO: add the code for cnetering a UserForm.

IN addition tot hose simple concners, you also have hte ability ot danmically create controsl on demand. This makes it possible to add/remove controls ot the USerForm as needed. This can be helpful if oyu want to create Control based on some proeprt yof the Worksheet but where you may not know how many times to do it in advance. For example. maybe you want to prvoide a LIstBOx with unque values for each column that was slected. IN advnace, you may not know the column coujtn so you need to create ListBoxes on demand. This can be done with UserForm programming.

TODO: example fo create a Control from scrathc
