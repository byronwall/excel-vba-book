## declaring the parameters (Subs and Functions)

When creating a new Sub or Function you are able to determine the inputs to your new creation. There are a handful of ways of handling the inputs:

- Put the inputs into the parameters of the Sub/Function and allow the caller to provide them
- Use knowledge of the spreadsheet to determine the inputs (or prompt the user for an input)

The main split here is: do you require the person typing the VBA to give you the inputs? Or, do you use some other approach like asking the user or just pulling the inputs from the spreadsheet.

The most common approach is to pull the inputs out of the spreadsheet. This seems counterintuive, but if you consider that the vast majority of VBA code is purpsoes wirtten for a single use, then it stands to reason that code will not be built on a large nujmber of Subs/Functiojns accepting parameters. The reason for this is that generally someone writes VBA to handle _their_ spreadsheet and so the VBA just reflects that spreadsheet. This works great for indivudal cases but can become a burden when building larger workflows. The main thing to consider for lager workflows is that as the complexity grows, tehre will be a large amount og code that is called multiple times or could be called separately from the main workflow. When thi sis the case, you are often served by pulling that code out into its own Sub/Functiojn wiiht aprameters.

To create a Sub or Function with parmaeters, you simply add them to the definition line:

```vb
Sub WithSomeName(firstParameter as String)

End Sub
```

This approach is very simple. You give the parameter a name and a type declaration. This is very nice because it nearly exactly matches the `Dim` statement with a Sub. That correspondence makes it very easy to start with an internally declared variable and then upgrade it to parmaeter. You can also go the other way: take a parmaeter and inline it into the Sub with some default or determined value. This is less common.

Once the parameter has been given a name and a type, you can simply use it within the Sub like any other variable. In this regard, your code will look the exact same. IF you are the person typing the VBA to use this Sub, then you will have to provide an appropriate variable as the parameter to make it all work.
