# SO item 062
I'm trying to change the font style of a chart through the selection.font.style property. Unfortunately this doesn't work, but I get an unsupported object or method error, this altough the documentation states that it should work

Documentation: [https://msdn.microsoft.com/en-us/library/aa213736(v=office.11).aspx](https://msdn.microsoft.com/en-us/library/aa213736(v=office.11).aspx)

Debug.print typename(selection) gives: ChartArea

The intelisense does not work either which complicates matters, what can be done?

My code

```
Selection.Font.style ="mystyle"

```

----

`Style` does not exist on `Font`. If you check the [documentation for that object](https://msdn.microsoft.com/en-us/library/aa174220(v=office.11).aspx), you will see that. Sometimes undocumented properties exist, but it is clear from trying that this is not one of those times.

Another indicator is that the `Home->Styles` part of the Ribbon is all greyed out once a Chart is selected

If you want to change the `Font`, you need to go through the properties available there: `Bold`, `Name`, etc.

You can apply a `ChartStyle` to the `Chart` (`Parent` of the `ChartArea`) which is the same as the items in the `Chart->Design` gallery in the Ribbon. Those are indexed by number and it is not obvious how those are determined. You can record a macro to get the desired number though.

Finally, a good idea for getting (some) help from Intellisense is to declare objects. In this case, `Font` does not exist on `ChartArea` which is not that helpful, but the properties are declared for `Font` when you hit the dot after it.

```
Dim cht_area As ChartArea
Set cht_area = Selection

'hitting the dot before Name brings up the list
'.. Font does not exist though
cht_area.Font.Name = "Arial"

```
