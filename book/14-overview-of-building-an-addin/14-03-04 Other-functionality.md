### Other functionality

THe other functionality that you can add is related to Events. You have great power when it comes to listening to events and triggering various actions. THe real difficulty is deciding what is an appropriate use of that power. Namely, when will you create an experience that benefits the user versus creating a very confusing workbook that is prone to breaking?

Before diving into what events can do, it's worth nting that potential downfalls of using them:

- They can be quite finicky sometimes. That is, using events adds a layer of complexity that tends to just complicate Excel and VBA. I don't have a technical explanation, but there seem to be a number of bugs that creep out of the dark once you start really using events.
- Your user can disable events at will and it can be quite difficult to determine when that was done. This is done with `Application.EnableEvents = False`.
- Events are triggered all the time for all sorts of reasons. If you are doing a lot of checking in Events, you will dramatically slow down the workbook.

With all of those warnings, there is nothing wrong with using Events. They generally do what you want and can be quite powerful. I add the caveats only because I have seen them ruin an otherwise working workbook. That complexity gets amped up a level when your Event code is inside an addin instead of the main workbook.

To really make the most of Events, you are going to need to use Class Modules. The reason is that your Events need to "latch on" to the host workbooks or worksheets, and the only way to do that is by using Class Modules. Normally, outside of an addin, you can simply open up the relevant VBA object (Workbook or Worksheet) and add the event code there. For an addin, you cannot add that code outside of the addin so you are in a bind. How then can you hook onto the Event? Fortunately, VBA makes this possible with the `With Events` command inside of a Class Module.

TODO: provide a concrete example of using this code
