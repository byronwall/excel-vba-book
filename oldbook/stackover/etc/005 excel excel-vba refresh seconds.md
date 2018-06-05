# SO item 005
I have a list of stock prices pulled from Google finance and placed in different sheets in my Excel. I'm wondering, Can I refresh Excel sheet every SECOND (not minute) according to the Google finance stock price?

----

This can be done without having a macro constantly running. It relies on the **Application.OnTime** method which allows an action to be scheduled out in the future. I have used this approach to force Excel to refresh data from external sources.

The code below is based nearly exclusively on the code at this link: [http://www.cpearson.com/excel/ontime.aspx](http://www.cpearson.com/excel/ontime.aspx)

The reference for Application.OnTime is at: [https://msdn.microsoft.com/en-us/library/office/ff196165.aspx](https://msdn.microsoft.com/en-us/library/office/ff196165.aspx)

```
Dim RunWhen As Date

Sub StartTimer()
    Dim secondsBetween As Integer
    secondsBetween = 1

    RunWhen = Now + TimeSerial(0, 0, secondsBetween)
    Application.OnTime EarliestTime:=RunWhen, Procedure:="CodeToRun", Schedule:=True
End Sub

Sub StopTimer()
    On Error Resume Next
    Application.OnTime EarliestTime:=RunWhen, Procedure:="CodeToRun", Schedule:=False
End Sub

Sub EntryPoint()

    'you can add other code here to determine when to start
    StartTimer

End Sub

Sub CodeToRun()

    'this is the "action" part
    [A1] = WorksheetFunction.RandBetween(0, 100)

    'be sure to call the start again if you want it to repeat
    StartTimer

End Sub

```

In this code, the StartTimer and StopTimer calls are used to manage the Timers. The EntryPoint code gets things started and CodeToRun includes the actual code to run. Note that to make it repeat, you call StartTimer within CodeToRun. This allows it to loop. You can stop the loop by calling the StopTimer or simply not calling StartTimer again. This can be done with some logic in CodeToRun.

I am simply putting a random number in A1 so that you can see it update.
