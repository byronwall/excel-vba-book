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
