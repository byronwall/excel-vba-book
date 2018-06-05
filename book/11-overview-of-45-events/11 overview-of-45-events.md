# overview of 4.5 events

chapter will focus on using events to interact with the user and also to drive more functional spreadsheets

the major events to focus on are the `Workbook` events, including:

* `SelectionChanged` which can be used to track when the user clicks on something (do this if clicking in this row)
* `Changed` which allows for watching cells and doing specific things if the change was somewhere specific
    * using the Intersect technique to determine if the change was in an area of interest
* disabling events while making changed during events
* Application.OnWait event to trigger something to take place at a given interval

other ways to interact with events include via class modules with the WithEvents designation.  These can be used to associate an object with an event and then wire up the code separate from the original macro code.  This section might be useful for charting events if I ever get that code put together

Other areas where events take place is via the Ribbon and also via different controls that can live on the sheet.  It would be good to discuss these as well.
