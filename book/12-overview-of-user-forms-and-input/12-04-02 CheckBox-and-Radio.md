### CheckBox and Radio

The CheckBox and Radio are cousins (or siblings?) of each other and will be dealt with at once. They allow for a Boolean selection of an option. For the Checkbox, you are allowed to indicate the on/off state of a given button. For a Radio, you are allowed to indicate the on/off state for a single option _within a group of options_. The main thing to note about the Radio is that by selecting one item, you will deselect the others. In this way, the uses of these two controls maps naturally to the tasks you are likely to see.

Aside from the Name, the main items to deal with are:

- Clicked event - just double click to get this one
- Value - note you get this by default using the name, but it will include a Boolean of the selected state
- Enabled - can be used to disable the control

That's about it. You can change the formatting and other stuff, but these items typically exist to get an input and get to the real work. They are very common when you are providing options to the user or otherwise want to direct downstream If/Switch statements.

Beware that the Click event may be changed multiple times depending on how it was triggered (TODO: is that right?).
