### EnableEvents

The more aggressive ooptiojn in this chapter is to disable events form firing. This has the ffect of inmprpving perfomrance bcause your code will be able to skip a number of potentially lsow steps. The downsdie f this approach ist hat someitmes you need events to fire in order to achieve a deisred result. This is espeically the case if you wrote the event handlers.

For events, there is one extra consderiong for when yu might disable them. If you possibly making changes to the Workbook or Workseeht in an event, you will liekly need to disable events whiel you make your changes. THe reaosn for this is to prevent an endless loop of your event handler processing the change you jut made. This is only relevant for a handful of event tpyes (Selection and Change are common) but it happens to be the case that this is a problem on teh most commonly used event handlers.

TODO: add an example of a full event handler that disables event hadnling
