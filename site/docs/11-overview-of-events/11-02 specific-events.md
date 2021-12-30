## specific events

For actual event handling code, it makes most sense to take a look at the specific events that can occur and show some techniques for handling them.

### Worksheet

THe Worksheet hsa a number of events which are commonly used. These include:

- Changed
- SelectionChanged
- Activate

These events roughly correspond to their name and are easy enough to handle. The idea with these is that you have a specific worksheet that you want to monitor for a specific event. In that case, you add the event using the VBE and then add the handling code.

The most common approaches for using these events is to track what the user is doing and then provide some additional functionality based on their actions. There are a number of reasons that you might want to respond to their input:

- Advanced usability where you allow the act of selecting a cell or cells to determine that some macro should run. You could imagine on a certain sheet that selecting a new cell may mean "please load more data about this row" and the VBA responds accordingly.
- Validation of user input. It is common to watch what the user is changing and then determine if that change is allowed or not based on specific rules.
- Starting a new action with some user input. I have previously used editing a cell to trigger a goal seek on that cell. This was quite nice because the VBA would undo my edit and then goal seek the previous cell to a new value based on its formula. This provided a very slick means of trigger goal seek without having to collect further user input.
- Refreshing some display. It's possible you set calculations to manual and then force a recalculate each itme the Worksheet comes into focus.

For the Worksheet, the typical flow is that you will create events at the Worksheet level only if you know than you will only want the code for a single Worksheet (or you are willing to duplicate it across Worksheets). If you want to have the same code run for _all_ worksheets, you should look at the Workbook events which provide better views of the entire Workbook.

For some examples here, I will show you how to do the goal seek business along with a separate event which watches user Selection and then processes the cells accordingly.

TODO: add the code and description for the Goal Seek event

TODO: add the code nad desc for some event which activates on Select

### Workbooks

For Workbooks, you have a lot of the same events as for a Worksheet. These events take the same parameters (TODO: is that right?) but allow you to watch for that event across all Worksheets in a Workbook. Depending on what you are watching for, this either makes perfect sense or is a real burden with false events that are not interesting. You will hav eot determine the proper scope for your events depending on what you need them to do. There are not fast rules here. The summary of possibly events then includes:

- Changed
- SelectionCHange
- Activated
- Opened
- BeforeSave
- Other Save events
- Closing

TODO: add some callback parameters above (and verify names)

TODO: review other events to see what may be useful

These events are similar to Worksheets except that they give you additional hooks that only make sense for a Workbook, specifically related to Saving, Opening, and Closing. There events can be quite useful if you want to do some amount of processing before the file is saved. One common approach I have taken is to delete extraneous data from a workflow spreadsheet to reduce the size and save time for a file. This can be used to great effect if your processing spreadsheet is generally pretty lean without the data it processes. You could also use this to delete a Chart that is large and having a big impact on file size. Once the file is opened, you can then use VBA to recreate the chart.

IN some cases, these events are used for that type of example where it seems like a lot of work to save some amount of hassle. Oftentimes, this is the case. You can spend a lot of time with event code to make it do exactly what you rwant. Sometimes for a user focused spreadsheet, however, this is the level of detail than tis required to ensure that everything will work every time for everyone.

TODO: add an example and desc for a data removal VBA

### Application

There are also a couple of events that exist at the Application level. These include:

- OnWait
- TODO: any others?

Application.OnWait can be used to trigger an event at some point in the future. This can then be used to trigger a block of code which runs at an interval by having the triggered code start a new event in the future. In this ways, you can use VBA to start a timer which executes every so often.

TODO: add the OnWait code for a timer

TODO: find another examples
