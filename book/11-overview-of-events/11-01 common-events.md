## common events

When talking about events, there are a couple of high level details to touch on:

- Where the events occur? That is, which object owns the event and how do you hook into it?
- When the does event occur?
- What are you allowed or not allowed to do while responding to the event?

For a spreadsheets, the events tend to occur within the objects of interest: Worksheet, Workbook, others (TODO: is that right?).

THe most common events are assoicated with the Workbook and Worksheet. If you want to tie into those events, you can typically just add a new handler using the VBE. Thsi process is actually fairly straightforward. The task bcomes more difficult when you want to tie into an event but you are not certain which object will fire the event, or you want to trakc an event that takes place outside of your code.

THe main considering when working wtih event handling code is that you need to be sensitive to teh fact that you can enter an endless loop if you accidentally trigger the same event as the one you are respondng to. This is suprprisingly easy to do if you are tied into the `Changed` or `Selection_Chnaged` events which trigger quite frequtnly.
