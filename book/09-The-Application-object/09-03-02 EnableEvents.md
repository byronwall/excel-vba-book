### EnableEvents

The more aggressive option in this chapter is to disable events form firing. This has the ffect of inmprpving performances because your code will be able to skip a number of potentially lsow steps. The downsdie f this approach is hat sometimes you need events to fire in order to achieve a desired result. This is especially the case if you wrote the event handlers.

For events, there is one extra considering for when yu might disable them. If you possibly making changes to the Workbook or Worksheet in an event, you will likely need to disable events whiel you make your changes. THe reason for this is to prevent an endless loop of your event handler processing the change you jut made. This is only relevant for a handful of event types (Selection and Change are common) but it happens to be the case that this is a problem on the most commonly used event handlers.

TODO: add an example of a full event handler that disables event handling
