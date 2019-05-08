## where Excel and VBA meet

The previous sections focused on the aspects of VBA that exist independent on Excel. It is worth ending this chapter with a section that discusses the general theme of where VBA and Excel actually do intersect. The main thing to remember is that VBA provides the programming constructs and language, while Excel exposes an object model to VBA that can be programmed against.

A good rule of thumb is that anything you can do in Excel can be done via VBA. This is probably not an exaggeration. The Excel object model is incredibly detailed and provides access to every nook and cranny of an Excel spreadsheet. This gives you enormous power to manipulate a spreadsheet in whatever way you can imagine, but it also means that it is easy to be overwhelmed by sheer volume of commands and objects that exist. Fortunately, there are only a handful of common/useful objects to start with and within those objects there is a significant amount of overlap. For example, the Range and Chart both expose formatting related properties (e.g. Border colors) but the ways of editing those are the same on both objects.

TODO: is that true about the Borders being the same?

TODO: finish this section... not sure where it's going
