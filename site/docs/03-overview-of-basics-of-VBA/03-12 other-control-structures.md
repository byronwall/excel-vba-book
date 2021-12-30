## other control structures

### With command

THe `With` command allows you to place a given variable within "scope" and avoid repeatedly typing that variable's name for each required call. The `With` command exists solely to reduce the number of times that a given object or variable name is typed. You are never required to ues a With command to accomplish a goal, but it can be helpful to clarify or avoid having too long of a code block. Having said that, a With block can be incredibly confusing to read especially when mixed with the always in scope function calls like `Range` or `Cells`. It is incredibly easy to avoid typing the required `.` to start a new line and accidentally refer to the globally scope object instead of your With scoped object. For this reason, I very rarely use the With command. When I do use it, I will typically only use it when I am working with a nested object that might be several levels deep. Having said that, I mostly avoid the With block by creating a variable which holds the object in question and using that instead. I have found that parsing a With block later can quickly become a confusing mess because of the difficulty of spotting the `.` which is critical.

If you read through some of the most common questions on the internet about "why my VBA no work?" you will quickly find issues with With blocks accidentally calling a globally scoped command. I have never asked those questions on the internet, but I have definitely been bitten by the same errors where a `.` is missed and the commang goes bonkers. It happens but is easily avoided by not using `With`.

### GoTo statements

`GoTo` statements are used to force execution to jump to a specific Label regardless of anything else that the program is doing. A `GoTo` statement is required for error handling but is otherwise frowned upon by programmers with experience in other languages. The problem is that a bad `GoTo` statement allows you to do much damage within a program because you can quickly corrupt your program state by jumping around. Also, other programming languages tend to include all of the nice features that have replaced places where `GoTo` was previously required. A good example of this is breaking out of a loop or skipping to the next item in a loop. The latter is typically handled with a `continue` statement in other languages. In VBA, this statement does not exist and you are required to use a `GoTo` if you want the functionality.

To make a GoTo statement work, you need to have a Label that the GoTo points to. An example looks like this:

```vb
Sub GoToExample()
    'doing some stuff

    If someConditiojn Then
        GoTo EndOfCode
    Else
        ' do some other stuff
    End if

EndOfCode:

End Sub
```

The rule for labels is that they are required to occur at the front of the lien (no indenting), they must be a single variable name without spaces, and they must end with a colon.

You should go to reasonable lengths to avoid using GoTo statements for anything other than error handling. They are the root of a lot of problems as execution order is concerned.

### Error Handling

One final control structure that exists is related to error handling. It is an inevitable consequence that computer programs will eventually throw errors. There are a lot of techniques and good practice that can avoid errors, but sometimes you will be forced to deal with an error. The alternative to error handling is usually a pop up that informs the user that something went wrong. For an experienced user, they may be able to handle the `Debug` or `Continue` or `End` decision, but your typical user will assume that your code has failed catastrophically. It's entirely possible that the error has no effect on your intended outcome, or that the error could be resolved if the user just hit `Continue` but the take home message is that _if_ something _has_ to happen to respond to an error (or a possibly error), then you need error handling.

The elements of error handling are simple:

- Determine when to allow an error to be thrown
- Determine what happens with execution when an error occurs
- Determine where to go back to once the error state has been addressed

The first decision to make is whether or not to allow errors to interrupt execution. By default, the answer here is "yes", an error will interrupt execution. If you want to handle this differently or reset it back to default, there are a pair of commands that can be used:

- `On Error Resume Next`, ignore all future errors, just keep trucking
- `On Error Goto 0`, stop execution immediately at the next error

If you are savvy about searching online for solutions to your problem, you will often see option 1 listed as the "go to" (or is it `GoTo`, ha!) solution for getting around an error. In the technical sense, yes, `On Error Resume Next` will absolutely get you around an error. It will by definition ignore the error and just keep going with execution. For the vast majority of workflows, this is an awful approach. Very often an error is indicating that something has gone awry from your expectations. If those expectations were reasonable, then it is very likely that future code will not work as intended. Therefore,e if you are getting an error, you should give serious consideration to finding the source of it before you `Resume Next` through it. Ignoring an error that should have been addressed nearly always causes more pain later.

THe other harsh approach to respond to an error is to force execution to stop immediately. This prompts the user with the popup about how to proceed. This prompt is helpful because it gives two options that may allow you to solve the problem. THe first is `Continue` which will attempt to run the line of code again that cause the issue. If the error still persists, then you will simply get it again. No harm. However, it is also possible to change the state of Excel while the prompt is visible. T his means that if your code was relying on an `ActiveChart` and you did not select one; you will be able to select a chart before hitting `Continue`. This can be a quick way out of a problem if you are confident where the error occurred. If you are programming only for yourself, this can also be a clean way around dealing with waiting for user input using another `GoTo` approach down below. Having said that, allowing a user to deal with an error prompt is absolutely awful in terms of usability.

The second way you can deal with these error prompts is by hitting `Debug`. This is likely the first response when an error occurs because you are very unlikely to know where the exact error occurs. Once you've seen it however, then you may be able to contue above. THe nice thing about debugging the error is that you get some powerful tools to try and solve the problem. For a full overview of debugging, check out the other section (TODO: add link). The specific features that are nice for dealing with error sinclude:

- Locals window, which will provide an overview of all the local variables and their current state
- Set next statement, which will allow you to skip over an error or rerun a line of code whose state may have changed between executions
- Immediate window, which will allow you to either run arbitrary commands or possibly output information about the program state.

All of those tools combined should make it possible for you to determine the source of an error. Once you have determine the source of an error, you can then set about resolving the error, again using the debug tools. Once you have solved the problem, you should give serious consideration to then adding that solution to the code using proper error handling techniques. Again, it is absolutely awful to present the user with an error dialog and expect them to be able to figure it out. Even if you are the user, you will absolutely tire of dealing with error prompts that can be handled with proper handling.

If you want to address an error, there are a couple of ways to handle that. They all rely on using the `On Error Goto LABEL` technique. This allows the code execution to jump to a specific place in your code. That area in your code is then able to do a couple of helpful things:

- Query the state of the `Err` object
- Attempt to address the error and then kick code back to the previous spot
- Provide the user with proper feedback before killing execution
- Log the issue accordingly before failing or prompting the user

With all of these approaches, the idea is simple: redirect execution to a known spot when the error has occurred. Once you are in a known spot, you can then step through possible problems and possible solutions. If you want, you are then able to send execution back to another spot to advance. IF you cannot resolve the error (or determine what caused it), you can then end execution all the same. Ideally you end execution with a better message than the normal prompt.

TODO: give an example of some error handling code

#### avoiding errors

Although this section is about error handling, the best error handling is an approach hta tmake is very difficult for an error to occur in the first place. As you call into specific VBA and Excel Subs, you will gain a feel for which ones can cause problems. On the VBA side, there are a number of specific calls that will lead to errors:

- Indexing into an array with a index that is not valid: `Sheets("SomeSheetThatIsMissing")`
- Attempting to use a property on an object that does not exist
- Sending invalid parameters to a function

All of those items above have the nice poprerty that you may be able to provide checks for when you will enter an error state. The upside of this approach is that you can use an `If...Then` statement to check for an error causing state and then step around it. Before using `Range.Value`, you can check that `If Not Range Is Nothing`. `Nothing` is the default value for a reference type before it has been set to a poper reference. You are always going to get an error if you attempt to use a `Nothing`. You can avoid a ton of errors being thrown by simply checking for Nothing and avoiding its use when it appears.

For a lot of arrays and other utterable objects, you have different approaches for checking inf something is a valid index before accessing it. For a `Dictionary`, there is the `Exists` method. For `Worksheets` and other Excel arrays, you are always able to iterate through all of the items to check for existing before then using the index. TODO: add example of iterating sheets. It is very rare for the performances of VBA to be affected by these types of checks. There ar instances where it is not appropriate, but in general, these techniques work fine.

#### Application.XXX functions

In some instances, it is possible to trade a runtime error for a return value that has a type of error. This occurs with the Application.XXX functions where XXX includes items in the list:

- Match
- TODO: any others?

This can be beneficial because when the function returns an error, you can then turn around and deal with it by checking `IsError`. If the function throws an error instead, you are forced to use proper error handling to catch the error and attempt to resume state.

#### common VBA errors

TODO: add section about 1004

TODO: add information about compile time errors vs. run time errors.

#### common Excel errors

In addition to the VBA errors, there are also a number of Excel specific errors that happen often enough that they should by addressed. Some of those common examples include:

- Using `ActiveXXX` without have `XXX` selected. This is most common with `ActiveChart` where it is possible to not have a Chart selected. This is not possibly with ActiveWorkbook or ActiveSheet since one will always be active. TODO: what about ActiveCEll?
- Using `Selection` when the "wrong" thing is selected. It is quite common to `Set` some vriable equal to Selection. If the wrong thing is selected, you will get an error about `Type Mismatch`
- Attempting to make a selection when it is not valid per the UI. This is most often the case when you attempt to Select a cell when its Parent Worksheet is not selected.
- Attempting to build a Range across Worksheets using `Union`
- Attempting to iterate through a Range of cells by checking `Range.Value` if the Range can contain errors. If this is possibl you will instead have to check for errors first.
- Attempting to access or change the `AutoFilter` if is has not been enabled first

There are also a ton of instances where some function returns `Nothing` and you do not check for it. T his most commonly occurs with:

- `Range.Find` where nothing was found
- `Intersect` where the two Ranges do not overlap
- TODO: add some others?

As a final note, it is owrht mentioning that the sign of a good programmer is one who has a feel for when errors can and cannot occur. You will begin to appreciate when it is needed to add error handling code versus when you know you will not need it. Too often as a beginner, you will be excluding error handling because you are unaware of what can go wrong. As you get better, you will start to exclude error handling because you actually know that no errors can occur. Until you get good, the result may look the same (no error handling code) but the result to the user is prompts and halted execution in one case.
