### common properties of a Point

The Point represents the lowest level when it comes to how the data and formatting of a Chart is built. In general, you do not have to actively go editing Points. This is because you will typically edit the appearance of the Series and the Axes to get the Chart that you want. There are however times when you get down to the metal and edit the properties of the individual points. Before describing how to do this, it may help to give an example or two for why you want to get down to this level:

- Delete a data point without touching the Series
- Add a DataLabel to the point if the value is below some threshold (or if some other Range has a value)
- Hide a Point from one series because you want it to show up in another one

Of the tasks above, only one of them (the second) has to be accomplished via the Points. The others _could_ be done via a different method, but you might find yourself in a spot where iterating some Points will save a ton of headache elsewhere. A cautionary note is that typically you should not be editing the properties of a Point; there is nearly always a better way to do these things. Part of the problem is that the settings you change will be quickly overwritten by changes in Excel or VBA. If you know you just need something done however, Points can be a quick way to make it happen.

TODO: look into ErrorBars here?

WHen thinking about working through the Points of a Series, consider the common properties you can change:

- HasLabel / DataLabel
- Value
- Formatting? (TODO: what are these)
- Hidden

TODO: finish this list

Note that in addition to the common properties, you can also change anything that can be changed from the normal Excel settings/properties window.
