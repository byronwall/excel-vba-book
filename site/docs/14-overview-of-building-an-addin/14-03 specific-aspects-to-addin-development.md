## specific aspects to addin development

Depending on the addin that you are creating, you may expect for it to have a handful of features available. In general, those types of features include keyboard shortcuts, special forms or user prompts, and possibly automatic features that fire depending on the user's action or the state of the workbook or Application.

### Keyboard Shortcuts

The simplest thing to do is to add keyboard shortcuts to your addin. There are two ways to do that:

- Open up the Macros form on the Developer tab. You can then hit "options" for a given Sub and assign a keyboard shortcut (TODO: add picture of this)
- That approach can sometimes be a pain to edit later, so you can also add code to your addin to add the shortcut.

The latter approach is nice because you can easily change the shortcut or the calling method. For addins, I will nearly always take the latter approach since it is much easier to deal with alter. For XLSM workbooks, I will do the former since it is easier to change from a workbook.

If you want to add the keyboard shortcut using code, use the code below. Ideally, you would put this in a Workbook_Open event that is called when the workbook opens. You can also use this approach to add/remove shortcuts depending on user input.

```vb
Public Sub SetUpKeyboardHooksForSelection()


    'SHIFT =    +
    'CTRL =     ^
    'ALT =      %

    'set up the keys for the selection mover
    Application.OnKey "^%{RIGHT}", "SelectionOffsetRight"
    Application.OnKey "^%{LEFT}", "SelectionOffsetLeft"
    Application.OnKey "^%{UP}", "SelectionOffsetUp"
    Application.OnKey "^%{DOWN}", "SelectionOffsetDown"

    'set up the keys for the indent level
    Application.OnKey "+^%{RIGHT}", "Formatting_IncreaseIndentLevel"
    Application.OnKey "+^%{LEFT}", "Formatting_DecreaseIndentLevel"

End Sub
```

### USer Forms

One of the nice features of an addin are adding custom forms to provide the user with a better experience. Creating a UserForm in VBA is dead simple, and this is the best bang for your buck in terms of creating a professional looking product. The simplest of forms with the simplest of features can save the end user hours and hours of time (I've seen it happen).

The nice thing here is that creating a UserForm in an addin is not any different than creating them normally. You simply create the UserForm. The only extra step is that you need to manage how/when the form is created and what information it has access to. Typically this is done by adding a button or using a keyboard shortcut. The only other issue is that you need to be aware of which Workbook or Worksheet is active when opening a UserForm if you are using ActiveSheet or ActiveWorkbook for anything. In general, inside an addin, you need to be careful with this commands since it is not always obvious that the ActiveXXX is the one you want to access.

### Helpful Commands

There are a couple of commands that exist outside of addins that become far more useful inside the addin. They are included below for reference:

- `ThisWorkbook` refers to the workbook that contains the code being executed. This is the surefire way to refer to the XLAM file that is running instead of the ActiveWorkbook. IN general, your addin will never be the ActiveWorkbook. This becomes relevant if your addin workbook contains sheets of data that may need to be accessed during runtime. You would use THisWorkbook to refer to those sheet.
- TODO: add any other commands that are addin specific

### Other functionality

THe other functionality that you can add is related to Events. You have great power when it comes to listening to events and triggering various actions. THe real difficulty is deciding what is an appropriate use of that power. Namely, when will you create an experience that benefits the user versus creating a very confusing workbook that is prone to breaking?

Before diving into what events can do, it's worth nting that potential downfalls of using them:

- They can be quite finicky sometimes. That is, using events adds a layer of complexity that tends to just complicate Excel and VBA. I don't have a technical explanation, but there seem to be a number of bugs that creep out of the dark once you start really using events.
- Your user can disable events at will and it can be quite difficult to determine when that was done. This is done with `Application.EnableEvents = False`.
- Events are triggered all the time for all sorts of reasons. If you are doing a lot of checking in Events, you will dramatically slow down the workbook.

With all of those warnings, there is nothing wrong with using Events. They generally do what you want and can be quite powerful. I add the caveats only because I have seen them ruin an otherwise working workbook. That complexity gets amped up a level when your Event code is inside an addin instead of the main workbook.

To really make the most of Events, you are going to need to use Class Modules. The reason is that your Events need to "latch on" to the host workbooks or worksheets, and the only way to do that is by using Class Modules. Normally, outside of an addin, you can simply open up the relevant VBA object (Workbook or Worksheet) and add the event code there. For an addin, you cannot add that code outside of the addin so you are in a bind. How then can you hook onto the Event? Fortunately, VBA makes this possible with the `With Events` command inside of a Class Module.

TODO: provide a concrete example of using this code
