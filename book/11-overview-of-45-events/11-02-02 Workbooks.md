### Workbooks

For Workbooks, you have a lot of the same events as for a WOrksheet. These events take the same parameters (TODO: is that right?) but allow you to watch for htat event across all Worksheets in a Workbook. Dpeending on what you are watching for, this either mkaes perfect sense or is a real burden wiht false events that are not interesting. You will hav eot determine the proper scope for your events dpeending on waht you need them to do. There are not fast rules here. The summar of possilby events then includes:

- Changed
- SelectionCHange
- Activated
- Opened
- BeforeSave
- Other Save events
- Closing

TODO: add some callback parametsrs above (and verify names)

TODO: review other events to see what may be useful

These events are similar to WOrksheets except that they give you additioanl hooks that only make sense for a Workbook, specifically realted to Saving, Opening, adn Closing. There events can be quite useful if you want to do some amount of processing befor ethe file is saved. One common approach I have taken is to delete extraneous data from a workflow spreadhseet to reduce the size and save time for a file. This can be used to great effect if your processing spreadhseet is generally pretty lean wihtout the data it processes. You could alos use this to delete a Chart that is large and having a big impact on file size. Once the file is opened, you can then use VBA to recreate the chart.

IN some cases, these events are used for that type of exmaple where it seems like a lot of work to save some amount of hassle. Oftentimes, this is the case. You can spend a lot of time with event code to make it do exactyl waht you rwant. Sometimes for a user focused spreadsheet, however, this is the level of detial tha tis requried to ensure that eveyrhting will work everytime for everyone.

TODO: add an example and desc for a data removal VBA
