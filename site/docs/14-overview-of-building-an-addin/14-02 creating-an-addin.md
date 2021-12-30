## creating an addin

Creating an addin is a relatively simple process. You start with a normal XLSM macro enabled file. From there, you save it as the add-in type (XLAM). That's it.

If you want to get more complicated, there is a property in the VBE that can be toggled to change the addin status. (TODO: add picture of that). You would only need to change that flag if for some reason you wanted to save something back to a normal XLSM workbook without changing the extension.

There is one additional process that can be done to change how the addin is created is that is if you are modifying the Ribbon for your addin. To do that, you will need to manually edit the XLAM file and change a file within it to add Ribbon support. You can do this manually or you can use a tool to help you out. Check the later section for details on that process.
