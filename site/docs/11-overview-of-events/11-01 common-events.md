## common events

When talking about events, there are a couple of high level details to touch on:

- Where the events occur? That is, which object owns the event and how do you hook into it?
- When the does event occur?
- What are you allowed or not allowed to do while responding to the event?

For a spreadsheets, the events tend to occur within the objects of interest: Worksheet, Workbook, others (TODO: is that right?).

THe most common events are associated with the Workbook and Worksheet. If you want to tie into those events, you can typically just add a new handler using the VBE. This process is actually fairly straightforward. The task becomes more difficult when you want to tie into an event but you are not certain which object will fire the event, or you want to track an event that takes place outside of your code.

THe main considering when working with event handling code is that you need to be sensitive to the fact that you can enter an endless loop if you accidentally trigger the same event as the one you are responding to. This is surprisingly easy to do if you are tied into the `Changed` or `Selection_Chnaged` events which trigger quite frequently.

### callbacks

ONe important point here is that all events are handled via callbacks. That is, you will create a Sub with a specific name and a specific signature which VBA then uses when the event occurs. This callback is important because it includes the information that you will need to discern what happened with the event. Each type of event can include its own specific parameters and your code can respond to them accordingly. This is important because if a cell was Changed, you will want different information than when a Worksheet was activated. In general, VBA is good about providing you useful lparmaeters so that your events can properly determine what took place. Despite the good parameters, you will very likely need to include If/Then code to determine if your event needs further processing.
