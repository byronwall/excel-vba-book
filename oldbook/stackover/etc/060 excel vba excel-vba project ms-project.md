# SO item 060
I have a collection as a global variable, that would have Project Task objects in them.

The structure of my macro would be the following:

```
Public TaskCollection As Collection
Sub Main()
   Set TaskCollection = New Collection

   GetData(List of project paths)

   For Each task in TaskCollection
        //ProcessTask()
   Next

End Sub

Function GetData(List of project paths)
    for each project path
         Open project p

            //do something else with the opened project...

            for each Task t in p.Tasks
                TaskCollection.Add t
            next
        Close project p
         //AFTER THIS, the TaskCollection object will be totally empty
   next
End Function

```

As i mentioned in the comments, after I close the project, from where I got the tasks into the TaskCollection, the TaskCollection loses it's values. The weird thing is, that it keeps the number of objects it had before, but they're all empty;

I tried to make a collection object locally in the GetData function, and then pass it in the TaskCollection global variable at the end, but the effect is the same.

----

You are adding references to objects when you add a `Task` to the `Collection`. These **references only have meaning so long as the objects they refer to exist**. Those objects are destroyed when the project is closed.

If you want to use their data, you will need to copy it using value types (`String`, `Integer`, etc.) and not refer to the objects. Or, you can keep the project open until you are done using the objects.
