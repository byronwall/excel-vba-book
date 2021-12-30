## more advanced events

This section will focus on using events in more advanced settings. In particular, the focus here will be on using Class Modules to allow for events to be attached to arbitrary objects that are not necessarily known at compile time. This is an advanced approach that is typically not required. Where it may be helpful is if you are building library code that needs to work in a range of settings. It may also be needed if you are trying to attach events to a Worksheet that will not exist until some other VBA has been run. In this case, you will attach to the same events as above, but you will add the event after the Worksheet has been created. It's worth noting for that specific example etha the Workbook can handle a large number of events on the Worksheet and will work for Worksheets that were created later.

TODO: add a section explaining how to use WithEvents

TODO: add some examples of attaching events to a new Worksheet.
