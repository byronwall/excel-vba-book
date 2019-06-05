### Event Handlers

Event Handlers are at the core of User Forms and making them useful. To be clear, your Form will do nothing without events. You could it to display static content from the designer mode, but it will do nothing useful. To make your Form become useful, you add Controls to it and then add Events to those Controls. Event Handlers are the glue (or wires) that take the actions perofmred on Controls and direct them somewhere useful. Evnets control everything from Clicking, Loading, Typing and everything else. Each COntrol has a unique set ofo events depending on what it can do, but in general, there's a bit of overalp between different controls.

To add an event handler, there are a couple of options:

- Double click on the Control in Design Mode, and oyu will get the default event handler created
- Go to the code view, and select the Control and then Event you want from the drop downs (TODO: add image)
- Type the Event handler based on the named of the Control and the event you want

If you know the default events, then option 1 is as good as the toehrs. IF you want to see a list of events beffore creating one, then you will go with optojn 2. You will pretty much never type the event handler out by handler unless you are copying it from somewhere else.

ONce you have created the ahndler, you simply add hte code that oyu want to fire in the event. One good tip here is to use the event handler to call other Subs. It's a good habit to not put logic or other execution based code into Even thNadlers. The reason for this is that you may want to perfomr the same action from multipl events. Putting the code in a handler makes it idfficult to resue the code because som ehadnlers have parameters nad other details that make it hard to arbitrary call them. Of course, I regularly put code into event Hnadlers, but at least I know I should avoid it. I am constantly reminded of why to avoid it when I have to extract code from one event to put into a SUb to call from another event.

One important note about Event handlers is that the hadnler can have some number of parmaeters that are included in the handler signature. These parameters ar etpyically used to pass along infomraiton related to the event. For ecample the key press event contains the key code of the key that was pressed. The Click event however has no paramteers. The presence of apramters is easy to check when the VBE creates the handler for an event since it will give the parameters.

TODO: given an example of using Handlers?

TODO: include a blurb about the Initialize event (if it was not addressed ealrier)
