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

The rule for labels is that they are required to occur at the front of the lien (no indenting), they must be a single variable name without sapces, and they must end with a colon.

You should go to reasonable lengths to avoid using GoTo statements for anything other than error handling. They are the root of a lot of problems as execution order is concerned.
