### ListBox

The ListBox is one control that has a number of options and a means of using it that are less obviosu than the other controls. It's a shame really that the ListBox is so unintuitive in VBA because it is quite powerful and other programming langugaegs have handled this better. THe idea behidn a ListBox is that it provides a list of items whose use can vary according to what you want. Some common applications include:

- Allow the user to select from one or multiple options in a list
- Provide some output to the user (and possinly then use that output as the input for next step)

THe input/output decision here is somewhat critical because the things that will annoy you about the ListBox break on this point. If you are collecting input, then really you have to also deal with output because at the end of the day, you have to put something in the ListBox in order for a user to select it. ONce you've handled the output stuff, then determinign which items have been sleected by the user is straightforward enough. Therefore, covering the output part is a good starting point.

To put items into the ListBox, you need to modify the List collection on the object. There are two ways to do this:

- Directly, via the List boject
- Indirectly, using the `AddItem` command

Either way you go, you have a couple of decisions after adding the item: what text do you want displayed for the item and do you want multiple columns? If you are dealing with a singel column, then you can simply add the text in the call for an addition and that's all. IF you are working with columns, then you will need to do two things:

- Set up the columns (using the editor or via commands) (TODO: add pictures or code here)
- Call the command to set the fields using the row and column nujmber (TODO: add some code)

ALthough I have described a simple process here, oftentimes, you will deal with something that is more complicated. THe issue comes when you want to maintain some reference to an obiject but you are reuqired to use a string for display purposes. This means that you need some means of maintaining that reference back to the object. There are options for dealing with this:

- Rely on the index of the objects matching (and not changing) and simply use the row index
- Create a Dictionary that stores the link between the string and the object
- Use some other object or Collection that can reference the object back to the string
- Serialize the object into the ListBox value (if multiple fields, join with a `|` or similar)

Each of those approaches has its pros and cons, btu the main idea is that you are often forced to deal with something that is typically much easier in other languages. My general approach is to rely on row index if I know that changes are not possible. This is common for a lot of code isnce yout ar elikely to contorl both side. If that is not ideal, then you can typucally find some way to store a reference ebetween the display value and the object using a Dictionary.

Once you hav eth einfomration in the ListBox, you can simply iterate the `Items` by index nad check the `Selected(index)` property to see if the item is sleected. Note that if you do not allow multiple selection, then you can also use the `SelectediNdex` property (TODO: is that right?).

TODO: add some code here to demonstrate iterate through a ListBox

Although this section has the most text, the ListBox is not always a pain to deal with. Typically they are much better than the alternatives (like using the Excel spreadsheet somehow) but require that you remember some boilerplate for accessing and changing items.
