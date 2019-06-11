### Workbooks

For Workbooks, you have a lot of the same events as for a WOrksheet. These events take the same parameters (TODO: is that right?) but allow you to watch for htat event across all Worksheets in a Workbook. Depending on what you are watching for, this either makes perfect sense or is a real burden with false events that are not interesting. You will hav eot determine the proper scope for your events depending on what you need them to do. There are not fast rules here. The summar of possibly events then includes:

- Changed
- SelectionCHange
- Activated
- Opened
- BeforeSave
- Other Save events
- Closing

TODO: add some callback parameters above (and verify names)

TODO: review other events to see what may be useful

These events are similar to WOrksheets except that they give you additional hooks that only make sense for a Workbook, specifically related to Saving, Opening, and Closing. There events can be quite useful if you want to do some amount of processing before the file is saved. One common approach I have taken is to delete extraneous data from a workflow spreadsheet to reduce the size and save time for a file. This can be used to great effect if your processing spreadsheet is generally pretty lean without the data it processes. You could alos use this to delete a Chart that is large and having a big impact on file size. Once the file is opened, you can then use VBA to recreate the chart.

IN some cases, these events are used for that type of example where it seems like a lot of work to save some amount of hassle. Oftentimes, this is the case. You can spend a lot of time with event code to make it do exactly what you rwant. Sometimes for a user focused spreadsheet, however, this is the level of detial than tis required to ensure that everything will work every time for everyone.

TODO: add an example and desc for a data removal VBA
