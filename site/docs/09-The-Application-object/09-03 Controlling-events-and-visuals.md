## Controlling events and visuals

THe previous section focused on controlling calculations, generally for the sake of performance. When seeking better performances, there are two other changes that are commonly made. They rely on disable the screen updating also disabling events. The former is a pretty tame change and is a no brainer if you want performance. There are very few downsides to disabling the screen. Disabling events can give a big boost in performances also, but there are a couple more risks involved. In addition to performance, there are other times where you need to disable events in order for your code to work.

The most common code for performances is repated here for clarity:

```vba
Application.CalculationMode = xlManual
Application.ScreenUpdating = False
Application.EnabledEvents = False

Application.EnabledEvents = True
Application.ScreenUpdating = True
Application.CalculationMode = xlAutomatic
```

What does that code do? Again, it forces calculation to manual mode, the screen to not update, and events to not fire.

### ScreenUpdating

Screen updating is one of those things that seems fairly silly if you are coming from another programming environment. In the vast majority of other programming settings, it is very uncommon for all of your changes to produce an immediate visual result. ON the one hand, this is doccifult and on the other, it is awful for performance. Exel takes the opposite approach since it is a user focused GUI that offers a scripting environment for uaotmatiojn. The default in Excel is that all of your commands will trigger the normal render and refresh that would have occurred had you manually made the change. IN practice, this can often produce a very cool effect where the commuter is quite literally doing all of the work for the user. once you get over the appeal of this, you will quickly be left with the question of: how much ƒaster will my code be if I do not process all of these visual updates? The answer: much faster.

There are very few risks to disabling the screen. The biggest risk is that you forget to enabled it again (usually because of an error) and then your usre will get odd behavior. In a lot of case, Excel will actually enable the screen anyways so the actual risk here is minimal.

The only real reason to leave the screen active is in the case where your automation can ataxy be usefully reviewed by the user. In that case, you may consider leaving the scren active so that the user can "watch" what is happening. Some users like to see what the code is doing.

### EnableEvents

The more aggressive option in this chapter is to disable events form firing. This has the ffect of inmprpving performances because your code will be able to skip a number of potentially lsow steps. The downside f this approach is hat sometimes you need events to fire in order to achieve a desired result. This is especially the case if you wrote the event handlers.

For events, there is one extra considering for when yu might disable them. If you possibly making changes to the Workbook or Worksheet in an event, you will likely need to disable events while you make your changes. THe reason for this is to prevent an endless loop of your event handler processing the change you jut made. This is only relevant for a handful of event types (Selection and Change are common) but it happens to be the case that this is a problem on the most commonly used event handlers.

TODO: add an example of a full event handler that disables event handling
