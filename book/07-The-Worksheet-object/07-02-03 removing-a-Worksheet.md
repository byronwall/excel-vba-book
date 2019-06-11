### removing a Worksheet

If you need to delete a Worksheet, it is a simple command again: `Worksheet.Delete`. The one downside to this command is that it will fire off a warning prompt if the Worksheet contained any data or was otherwise not "blank". This warning box will stall the execution of your VBA until it is addressed. This is a major issue for any serious workflow since your users will have to constantly click "Yes" to delete the Worksheet but they may also have no idea what they are deleting. To avoid this issue, you will nearly ALWAYS wrap the `Delete` command with the commands to disable and then re-enable the alerts. The typical code looks like:

```vb
Application.DisplayAlerts = False
Worksheet.Delete
Application.DisplayAlerts = True
```

When doing this dance, be absolutely certain that you re-enable the alerts. Excel will not do it for you. You may benefit from creating a new helper Sub which contains the above code as a `DeleteSheet` command to avoid constantly adding those alerts.

TODO: add a note about when to create a new Worksheet vs. a new Workbook and the pros/cons there (maybe put this in the workflow section of book)
