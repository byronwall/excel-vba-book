## common patterns

There are a number of patterns that are very common with Events. These patterns typically exist to avoid causing a problem or to avoid extra work where possible. Most VBA is not performance critical, but it is possible for an event to be called hundreds of times for a given chucnk of code. Since this is true, you can start to have an immediate impact on performance if your event handling code includes a number of unnecessary steps. As a side note, this is a good reminder that when trying to speed up code, you will nearly always do better to add `Application.EnableEvents = False` before your performance critical code; this assumes that your VBA does not rely on events firing to function properly.

### Intersect

The first is the `Intersect` technique to determine if a Range that was affected by an event was a Range of interest. With this approach, you define a Range which includes your "interesting" cells. You then do a `If Not Intersect(rngEvent, rngTarget) Is Nothing` to see if the intersection of the callback Range and the desired Range overlap. If they overlap, yhen you typically execute some code. This allows you to quickly filter out Ranges which have changed but are not relevant to Athena code you need to run.

TODO: add a code sample here

### Application.EnableEvents = False

One of the biggest gotchas with Events is that you can quickly and accidentally create an endless loop of Event code running if your event handler is able to retirgger the original event. This is quite common if you are looking at the Selection and then change the selected cell. The same can happen if you are using an event to watch for a change and then you respond with additional changes. Both of these accidents are so common, that you should seriously consider always disabling events in your handler. It is quite rare that you will need an other event to fire following your own processing.

The main thing to remember here is that you really need to enable events again. Excel will not do this for you. You can create odd situations if you have an error in your code that goes unchecked. This situation can mean that events are disabled. For really sensitive, user focused code, you should add a proper error handler and enable events following that.

To handle this event, the code is quite simple:

```vb
Sub EventHandler()
    'disable events
    Application.EnableEvents = False

    '' do some stuff

    're-enable events
    Application.EnableEvents = True
End Sub
```
