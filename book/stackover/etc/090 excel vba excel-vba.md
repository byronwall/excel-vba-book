# SO item 090
We are developing a macro to be used by customer organizations. Is it possible to get rid of the macro security warning messages by signing the macro file with a Digital signature?

We already have a code signing certificate from Comodo but when I opened the signed macro file in another computer (an old mac) it still gave the warnings. Was it because we didn't do something right or because it is not possible to get rid of the warning at all.

----

It still hinges on the security settings on the computer which is opening the code. There is an option in there `Disable all macros except digitally signed macros` which must be selected for the signature to matter.

![settings](https://i.stack.imgur.com/iGWeh.png)

My current settings would still prompt for your add-in.
