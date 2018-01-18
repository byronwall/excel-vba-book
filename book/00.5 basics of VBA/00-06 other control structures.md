## other control structures

### With command

THe `With` command allows you to place a given variable within "scope" and avoid repeatedly typing that variable's name for each required call.  The `With` command exists solely to reduce the nujmber of times that a givne object or variable name is typed.  You are never required to ues a With command to accomplish a goal, but it can be helpful to clarify or avoid having too long of a code block.  Having said that, a With block can be incredibly confusing to read especially when mixed with the always in scope function calls like `Range` or `Cells`.  It is incredibly easy to avoid typing the required `.` to start a new line and accidentally refer to the globally scope object instead of your With scoped object.  For this reason, I very rarely use the With command. When I do use it, I will tpyically only use it when I am workign with a nested object that might be several levles deep.  Having said that, I mostly avoid the With block by creating a variable which holds the object in question adn using that instead.  I have found that parsing a With block later can quickly become a confusing mess becuase of the difficulty of spotting the `.` which is critical.

If you read through some of the most common questons on teh interent about "why my VBA no work?" you will quickly find issues with With blocks accidentally calling a globally scoped command.  I have never asked those questions on the internet, but I have definitely been bittne by teh same errors where a `.` is missed and the commang goes bonkers.  It happens but is easily avoided by not using `With`.

### GoTo staements

`GoTo` statements are used ot force execution to jump to a speciifc Label regardless of anythign else that the progrma is doing.  A `GoTo` statement is requried for error handling but is otherwise frowned upon by programmers with expereince in other languages.  The rpoblem is that a bad `GoTo` statement allows you to do much damage within a program because you can quickly corrupt your program state by jumping around.  Also, other programming languages tend to include all fo the nice features that have replaced places where `GoTo` was prevously required.  A good example of this is breaking out of a loop or skipping ot the next item in a loop.  The latter is tpyically handled with a `continue` statement in other langugaes.  In VBA, this statemnet does not exist and you are reuqired to use a `GoTo` if you want the functionality.

To make a GoTo statement work, you need ot have a Label that the GoTo points to.  An example looks like this:

```vb
Sub GoToExample()
    'doing some stuff

    If someConditiojn Then
        GoTo EndOfCode
    Else
        ' do some other sutff
    End if

EndOfCode:

End Sub
```

The rule for labels is that they are reuqired to occur at the front of the lien (no indenting), they must be a single vairable name without sapces, and they must end with a colon.

You should go to reasonable lengths to avoid using GoTo statements for anything other than error handling.  They are the root of a lot of problems as execution order is concerned.

### Error Handling

TODO: add some content about Error handling
