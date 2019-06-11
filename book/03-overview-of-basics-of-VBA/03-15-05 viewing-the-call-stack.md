### viewing the call stack

One final feature which is useful is to check the Call Stack. The Call Stack is a list of all the proceduring Subs or Functions that are "active" preceding the current command. It gives you a list of all the places that came before your current line of code. The Call Stack is invaluable when you have started debuggin following an error because oftentimes you will not know how you reached a given spot. This is especially true if you are debugging code that is used in multiple places.

To see the Call Stack do View->Call Stack. You can then double click on an item and jump back to that spot. Note that the VBE will attempt to show you the vales of variables at that location which can be very helpful.

The Call Stack can be very helpful if you are using recursive code that calls itself. This code can be very hard to debug because oftentimes a breakpoint will trigger more than you want. IF you are waiting for an error on the 8th time through a FUnctiojn, then you don't want to skip the breakpoint 7 times. Instead, you can wait for the error, then use the Call Stack to step back through the previosu iterations and see what happened.

TODO: add a picture of the call stack
