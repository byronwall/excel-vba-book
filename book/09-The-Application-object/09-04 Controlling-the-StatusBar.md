## Controlling the StatusBar

Did you know you can control the STatusBar? Did you know that the area at the bottom of you screen is called the Status Bar? This is an area that can be used to provide feedback to the user. It can be quite helpful for a long running calculation where you intentionally disable all of the normal feedback than the user receives (screen updates, events, etc.). If you have done that, be aware that you can still provider an updating message to the user.

```vba
Application.StatusBar = "Some message"
```

This functionality is best used when you have a measurable way of providing progress feedback. This is commonly one when you are looping through a large list of item sand processing each one in turn. Depending on how quickly you can process a single item, you may choose to update the StatusBar to provider the progress to the user. Biggest issue to be aware of is that you can overload the STatusBar and create a situation where Excel is slow processing all of your STatusBra updates. If this is your problem, you can usually remedy this with a quick modulo function to only update the status every 10th iteration or similar.

TODO: add the general purpose status tracking code
