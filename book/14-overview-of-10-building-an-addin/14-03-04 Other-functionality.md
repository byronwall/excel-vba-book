### Other functionality

THe other functiojnality that you can add is related to Events.  You have great power when it comes to listening to evnests and tirggering various actions.  THe real difficulty is deciding what is an appropraite use of that power.  Namely, when will you create an experiecne that benefits the user versus creating a very confusing workbook that is prone to breaking?

Before diving into what events can do, it's worht nting that potential downfalls of using them:

- They can be quite finicky somtimes.  That is, using events adds a layer of complexity that tends to just complicate Excel adn VBA.  I don't have a technical explanation, but there seem to be a number of bugs that creep out of the dark once you start really using events.
- Your user can disbale events at will and it can be quite difficult to determine when that was doen.  THis is done with `Application.EnableEvents = False`.
- Events are triggered all teh time for all sorts of reasons.  If you are doing a lot of checking in Events, you will dramatically slow down the workbook.

With all of those warngins, there is ntohign wrong with using Events.  They generally do waht you want and can be quite powerful.  I add the caveatas only because I have seen them ruin an otherwise working workbook.  That cmoplexity gets amped up a level when your Event code is inside an addin instead of the main workbook.

To really make the most of Events, you are going to need to use Class Modules.  The reason is that your Events need to "latch on" to the host workbooks or worksheets, and teh only way to do that is by using Class Modules.  Normally, outside of an addin, you can simply open up the relevant VBA object (Workbook or Worksheet) and add the event code there. For an addin, you cannot add that code outside of the addin so you are in a bind.  How then can you hook onto the Event?  Foretunatley, VBA makes thsi possible iwth teh `With Events` command inside of a Class MOdule.

TODO: provide a concreate example of usign this code
