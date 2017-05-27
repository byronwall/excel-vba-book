## introduction to creating an addin

This chapter will focus on creating an addin for Excel using VBA.  There are other ways to create an addin but using VBA is simple becuase it can be done entirely from Excel and the Visual Basic Editor.  The main distinction between an addin and other VBA code is that an addin is meant to be avialable to all open Workbooks without having to put the code insdie a Workbook.  This can be a very nice thing to have if you are regularly do the same or similar operations across different Workbooks.  The alterantive to an addin is often to maintain a lirbary of code that you regualrly export/import into macro enabled files as needed.  This can create a mess as you cahgne code in one file but not in another.  The alternative also typically requires you to put the code inside a the Workbook adn make it macro enabled.  For certain applications, this is a non-starter.  The one otehr altnertive to a true addin is to create a Workbook that contains the code you want, adn then you can open thaty fiel and execute the code in the context of whatever other files are open.  This works, and creating an addin can be viewed as the logical condlusion of this approach.  More than the logical conclusion, this is actualyl teh first step for creating an addin.

When considering whetehr or not to create a proper addin with your code, consider the following:

* An addin provides a nice package for helper code and UDFs that might be used in multiple places
* An addin has easy access to teh Ribbon and can create its own Ribbon tab
* An addin can be put in a central location and used as a repository of code for an organization (works best if the file is read-only)

Item 1 in the list above is typically enough of a reason to consider creating an addin.  A common example of an addin is as a personal repository of VBA code.  This typically replaces the use of the Personal Workbook, which I have never found to work well.

When considering a personal addin, one of the biggest upsides is that you can always open the VBE and have immediate access to your library of code. This makes it easy to make edits and save the new addin.  Immediately, your updated code is avialable for future use in all your Workbooks.

There are a couple of downsides related to addins:

* UDFs from an addin require that anyone opening the spreadsheet has the addin loaded
* For code in a single Workbook, it is often easier to simply use a macro enabled Workbook and save the code directly there
* Some folks are highly resistnat to "installing an addin" but will happily open a XLSM file.  These are equivlanet in the case of opening an addin, but the hesitation still exists.

Point 2 above is worth expadning on.  Sometimes it's tempting to add code to an existing addin that make sesne only in the context of a single file.  This works well if you and everyonoe else have the addin.  This starts to become a nuisance when you are constnatly going through your addin to find code that should have been place in a Workbook to start. The cleaner way to store cdoe that may be useful later is to place a copy of it in a perosnal addin.  This ensures taht the orignal code is always avialable in teh Workbook adn that future updates to teh code don't break the orignail applicaiton.
