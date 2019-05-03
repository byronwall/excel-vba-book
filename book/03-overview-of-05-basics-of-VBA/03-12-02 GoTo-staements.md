### GoTo staements

`GoTo` statements are used ot force execution to jump to a speciifc Label regardless of anythign else that the progrma is doing. A `GoTo` statement is requried for error handling but is otherwise frowned upon by programmers with expereince in other languages. The rpoblem is that a bad `GoTo` statement allows you to do much damage within a program because you can quickly corrupt your program state by jumping around. Also, other programming languages tend to include all fo the nice features that have replaced places where `GoTo` was prevously required. A good example of this is breaking out of a loop or skipping ot the next item in a loop. The latter is tpyically handled with a `continue` statement in other langugaes. In VBA, this statemnet does not exist and you are reuqired to use a `GoTo` if you want the functionality.

To make a GoTo statement work, you need ot have a Label that the GoTo points to. An example looks like this:

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

You should go to reasonable lengths to avoid using GoTo statements for anything other than error handling. They are the root of a lot of problems as execution order is concerned.
