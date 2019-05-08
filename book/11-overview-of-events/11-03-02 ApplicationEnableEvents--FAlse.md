### Application.EnableEvents = FAlse

One of the biggest gotchas with Events is that oyu can quickly and accidentally create an endless loop of Event code running if your event handler is able to retirgger the original event. This is quite common if oyu are looking at the Selection and then change the selected cell. The same can happen if you are using an event to watch for a change and then you rrespond wiht additiojnal cahnges. Both of these accidents are so common, that you should seriously consider alwyas disbaling events in your handler. It is quite rare that you will need an otehr event to fire following your own processing.

The main thing ot remember here is that you really need ot enable events again. Excel will not do this for you. You can create odd situatiojns if you have an error in your code that goes unchecked. This sitautiojn can mean that events are disabled. For really sensitive, user focused code, you should add a proper error handler adn enable events following that.

To handle this event, the code is quite simple:

```vb
Sub EventHandler()
    'disable events
    Application.EnableEvents = False

    '' do some stuff

    'reenable events
    Application.EnableEvents = True
End Sub
```
