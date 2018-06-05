### Worksheet

THe Worksheet hsa a number of events which are commonly used.  These include:

* Changed
* SelectionChanged
* Activate

These events roughly correspond to their name and are easy enough to handle.  The idea wiht these is that you ahve a specifc cworksheet that you want to monitor for a speciifc event.  In that case, you add the event using the VBE and then add the handling code.

The most common approaches for using these events is to trakc what the user is doing and then provide some adidtioanl functionality based on their actions.  There are a number of reasons that you might want to respond to their input:

* Advanced usability where you allow the act of sleecting a cell or cells to determine that some macro should run.  You could imagine on a certain sheet that sleectng a new cell may mean "please load more data baout this row" and the VBA responds accordingly.
* Validatjon fo user input.  It is common to watch what hte user is changing and then determine if that change is allowed or not based on specific rules.
* Starting a new action with some user input.  I have previously used editing a cell to trigger a goal seek on that cell.  THis was quite nice because the VBA would undo my edit and then goal seek teh previous cell to a new value based on its formula.  This provided a very slick means of trigger goal seek without having to collect furhter user input.
* Refershing some display.  It's possible you set calcualtiojns to manual and hten force a reclacualte each itme the Worksheet comes into focus.

For teh Worksheet, the typical flow is that you will create events at the WOrksheet level only if oyu know tha tyou will only wnat the code for a single Worksheet (or you are willing to duplicate it across Worksheets).  If you want to have the same code run for *all* worksheets, you should look at the Workbook events which provide better views of the entire Workbook.

For some examples here, I will show you how to do the goal seek business along with a separate event which watches user Selection adn then processes the cells accordingly.

TODO: add the code and description for the Goal Seek event

TODO: add the code nad desc for some event which activates on Sleect
