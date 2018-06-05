# SO item 085
I have an Excel AddIn (`.xlam` file) and within it is a few macros and my attemt at a custom ribbon tab. The Macros work as expected but now I am trying to make a ribbon to call them to be more user friendly. I have the ribbon and a button which works, and a dropdown menu which I cannot figure out. I am unsure of what the Macro Parameters need to be. Below is what I have thus far.

The XML for the ribbon (Works!):

```
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="My_Tab" label="My Tab">
        <group id="NC_Material" label="Button1">
          <button id="Button1" label="1st button" size="large" onAction="Module1.Button1Click" imageMso="ResultsPaneStartFindAndReplace" />
          <dropDown id="DropDown" onAction="Module1.DropDownAction">
            <item id="ddmItem1" label="Item1" />
            <item id="ddmItem2" label="Item2" />
          </dropDown>
        </group >
      </tab>
    </tabs>
  </ribbon>
</customUI>

```

The VBA in Module1 of the `.xlam` files VBA Project:

```
'this one works
Public Sub Button1Click(ByVal control As IRibbonControl)
    call Button1Macro
End Sub

'this one does not work
Public Sub DropDownAction(ByVal control As IRibbonControl)
    call DropDownMacro
End Sub

```

I am getting errors when I change the value of the drop down menu in my ribbon. I do not know what parameters I need for the onAction macro of the drop down menu. I have been unable to find a good reference or example.

I am unable to do this using Visual Studio and cannot download any utilities or other programs.

Thanks in advance.

----

If you want to save your sanity, you can use [Andy Pope's Ribbon Editor](http://www.andypope.info/vba/ribboneditor.htm) which will generate the callbacks automatically (at least the text for them). Then you just copy them into the VBA module. (I am not affiliated with that add-in, I just use it and recommend it highly)

Using that tool, I get the following for `onAction` within a `Dropdown`

```
Public Sub DropDownAction(control as IRibbonControl, id as String, index as Integer)
'
' Code for onAction callback. Ribbon control dropDown
'

End Sub

```

Working code is also shown [in this answer](http://stackoverflow.com/questions/4562550/getting-the-selected-item-from-the-dropdown-in-a-ribbon-in-word-2007-using-macro) which was near the top of my search.
