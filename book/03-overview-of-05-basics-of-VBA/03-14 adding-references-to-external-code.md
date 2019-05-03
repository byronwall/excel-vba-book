## adding references to external code

This section will cover how to add References to other files and programming components. There are 2 main reasons for why you would need to do this:

- You want to access some code from an Excel file that you or someone else created
- You want to access code from an existing component on your computer

The latter reason on the list is the more common reason for adding a Reference. There are a handful of common references that are added if you want additional components that are not available by default. Of these, the most common include the Microsoft Scripting Runtime and the references to other Office programs. For example, if you want to create a Dictionary, you will need to reference the Scripting Runtime. In general, there are a number of references that are nearly guaranteed to exist on all Windows computers. Having said that, there are also a handful of references that are commonly made where the required file may not be available. This uncertainty about the files available on a system is the major downside of using these references.

TODO: add a list with other common references and what they might include

For the first item, there are times where you have created some code that would be useful to use somewhere else but that you don't want to copy. This can be common for using helper code that you know is included in another file. The major drawback to this approach is that you are creating a permanent link between the file and the reference. This means that the one file will quit working if the reference ever moves or becomes unavailable. Despite this drawback, there are times where this can be convenient and the drawbacks less significant.

To add a reference is relatively simple. You simply go to Tools -> References. You can then check the boxes for any references that you would like to add. To add a reference to an existing Excel file, you will have to browse to the file and select it that way.

TODO: add some images of how to add a component
