# SO item 018
I am looking for the best way to deploy Excel Macros to users. My goal is to make it super easy for end users to install and promote use by adding to the addin toolbar. I know that there are a number of help articles on this topic but couldn't find anything that covered this exact issue. Can you please help and excuse me if this is a noobie question. Please see below for replication steps for my issue.

1.  I have added the code below as a worksheet event on "This Worksheet" of an excel macro file
2.  I add the main code to a module that it references
3.  I save this as an .XLAM in the addin roaming folder
4.  I enable this as an addin in EXCEL 2013
5.  After I install it adds the button to an add in tab
6.  It works until I close Excel in which case the button disappears
7.  It is still under active add ins but not on the toolbar

The code:

```
Option Explicit

Dim cControl As CommandBarButton

Private Sub Workbook_AddinInstall()

On Error Resume Next 'Just in case

'Delete any existing menu item that may have been left.
Application.CommandBars("Worksheet Menu Bar").Controls("Super Code").Delete

'Add the new menu item and Set a CommandBarButton Variable to it
Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add

'Work with the Variable
    With cControl
        .Caption = "Convert Survey Reporter Tables"
        .Style = msoButtonCaption
        .OnAction = "CMB_General_Table_Formatting"
        'Macro stored in a Standard Module
    End With

On Error GoTo 0
End Sub

Private Sub Workbook_AddinUninstall()
    On Error Resume Next 'In case it has already gone.

   Application.CommandBars("Worksheet Menu Bar").Controls("Convert Survey Reporter Tables").Delete
  On Error GoTo 0

End Sub

```

----

If you want an alternative to using VBA to build the interface, I have previously deployed Excel add-ins (XLAM files) using some variety of Ribbon XML. This allows for very fine-grained control of the resulting interface and does not require you to work in VBA to build the interface. For most applications, I have found it is much easier to build the Ribbon components outside of VBA and then wire up the callbacks in VBA.

For the end user, I think this approach also delivers a better looking add-in since the resulting interface has its own Ribbon tab (or you can add to any of the existing ones) instead of being in the Add-ins Ribbon tab.

If you want to pursue this route, I highly recommend using the [Ribbon X Visual Designer](http://www.andypope.info/vba/ribboneditor.htm) to build the interface and set callbacks. I have used it to build an add-in that had more than 50+ features accessible by buttons and other Ribbon form controls. It was fairly painless once I got going.
