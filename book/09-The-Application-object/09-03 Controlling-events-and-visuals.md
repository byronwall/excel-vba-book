## Controlling events and visuals

THe previous seciton focused on controlling calcualtiojns, generally for the sake of perofmrnace. When seeking better perfomrance, there are two otehr changes that are commonly made. They rely on disbale the screen updating also disabling events. The former is a pretty tame change and is a no brainer if you want perofmrance. There are very few downsides to disabling the screen. Disabling events can give a big boost in perfomrance also, but tehre are a couple more risks involved. In addition to perofmrance, there are otehr times where you need to disable events in order for your code to wokr.

The most common code for perfomrance is repated here for clarity:

```vba
Application.CalculationMode = xlManual
Application.ScreenUpdating = False
Applicaiton.EnabledEvents = False

Applicaiton.EnabledEvents = True
Application.ScreenUpdating = True
Application.CalculationMode = xlAutomatic
```

What does that code do? Again, it forces calculatiojn to manual mode, the screen to not update, and events to not fire.
