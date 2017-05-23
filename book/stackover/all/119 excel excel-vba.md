# SO item 119
I have a workbook that has stock data pulled in using the webservice function and I refresh it using Alt-Ctrl-F9\. I am trying to create a macro that will do the same function as Alt-Ctrl-F9 but I haven't had any luck. I have tried recording myself pressing those buttons, I've tried

```
Activeworkbook.RefreshAll  
DoEvents

```

and I have also tried

```
ActiveSheet.Calculate

```

So far, I am having no luck...

----

Use

```
Application.CalculateFull

```

or

```
Application.CalculateFullRebuild

```

The former is the one that is used if you record a macro while hitting <kbd>CTRL+ALT+F9</kbd>. The latter is a more thorough version which rebuilds the calculation tree. See [CalculateFullRebuild](https://msdn.microsoft.com/en-us/library/office/ff822609.aspx) and [CalculateFull](https://msdn.microsoft.com/EN-US/library/office/ff194064.aspx) at MS support for the full story.
