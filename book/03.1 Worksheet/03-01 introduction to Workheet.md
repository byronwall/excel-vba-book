# The Worksheet object

## introduction to the Worksheet object

This chapter will focus on the aspects of the Worksheet that appear commonly in VBA code.  This chapter is a little shorter than otehrs because in general, the Worksheet is a conduit to more useful things.  There is very little that takes place within the Worksheet object taht is not just a pass through to the more interesting details (e.g. Range or Chart). Having said that, there are a handful of areas that are relevant to teh Worksheet and not accessible anywhere else.  Those specific areas include:

* Creating and managing Worksheets -- this sounds obvious but managing the referneces to Worksheets becomes a major issue when workign with large, complicated workflows
* Print layout, printing, and exporting
* Locking and setting passwords on Worksheets
* Managing the properties of the Worksheet itself including Name, tab color, etc.

TODO: any other Worksheet things?

Of hte topics listed above, the most important area is actually creating and managing the Worksheets in a complciated workflow.  This is closely related to working with Ranges since presumably you create the Worksheet to put data into or something else into it.  Managing the referneces to Worksheets can be a big deal and determining how best to access or select a given Worksheet can be important.  In addition to getting references, there are a handful fo times where you actually need to Activate a WOrksheet. Knowing when this is and is not requried is important.

TODO: when do you have to Activate?
