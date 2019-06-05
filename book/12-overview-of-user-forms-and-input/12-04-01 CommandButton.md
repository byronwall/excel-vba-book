### CommandButton

The CommnadButton or simply button is one of the most common controls to use. Its use is simple, known to everyone, and easy enough to program against. A button does one thing: get clicked. The event you want to know about is the `_Clicked` event. Fortunately, the VBE will atuomtallica create and wire up this event if you double click the button on the Designer versiojn of the form. This makes it dead simple to create the button code that you want: just double click the button.

Note that the default event will be created with the current name of the control. To avoid this, you need to change the name of the button before you create the event. Be aware that VBA and the VBE are not that smart with respect to naming things nad wiring up changes. If you change the name of the button after you create the event, your event will not work. You should not chagne the button name after creating the vent (or plan to recreate it).

Other properties of the button that might be used:

- Value - will change the text that is displayed (TODO: is this right?)
- Enabled - will change wheteher the button can be pressed and will change the visuals. Useful when you want to show that an option could be possible but is not currently allowed or enabeld
- Formatting and other visuals - you may change this on the proeprty editor but it is far less common to modify the fomratting once you are running. It can be done but is not common.

That's it; buttons are simple.
