### Other Controls

There are a couple of other controls that you may see that are summarized here:

* Label: these don't do much other htan provide some fixed text when could be changed later (I rarely ever do it)
* RefEdit: this contorl tehcnically allows you to select a Range from Excel.  They are quite buggy.  Depending on you main goal, you may od much better to use `Application.InputBox(Type:=8)` to access a Range.
* Tabs: these can be heplful for organizing a complicated workflow.  You will find yourself wanting to change the active tab and possibly limit access to later tbas.
* Wells?, whatever it's called, there allow you to Grouo controls. These may be requried for a Radio to work like you want (if you have mutliple sets of Radios on a singel form).
* 