## Introduction to VBA

This chapter will focus on the basics of VBA that are essential to using VBA to work with Excel. The upside of VBA is that it has a very simple instruction set. The downside of VBA is that it has a very simple instruction set. Fortunately, the vast majority of Excel/VBA interaction can be handled with very simple instructions. **The main difficulty with using Excel/VBA is not the VBA side of things, it's managing the object model for Excel**. This object model does not introduce new commands, but it does add a large number of interrelated objects, properties, and Functions that need to be known at some level to do anything. If you compare the length of this chapter to the length of the book, you will get a sense of what is meant by this.

An important thing to remember about VBA is that it exists outside of Excel, in some sense. VBA (Visual Basic for Applications) is derived from VB6 which is a legitimate programming language that (previously) was used for serious programming. These days (ca. 2019), no one starts a new project looking to use VB6; it doesn't offer the features of modern programming languages. That VBA exists outside of Excel means that there are certain parts of the language that are independent of anything Excel has to offer. These aspects of VBA are the core parts of the language, and you need to understand them before you can do anything related to Excel.

Having said all of that, VBA consists of several key instructions:

- Declaring and setting variables
- Declaring and calling Subs and Functions
- Logic structures
- Loop structures
- Other control structures (Errors and Goto)

In addition to those aspects of using the language, there are a handful of details related to programming in general that are worth hitting:

- VBA 101, opening the VBE and getting started
- Adding references (how and why)
- Debugging code and using the tools provided

The flow of this chapter will hit on the VBA 101 question first. From there, we'll hit the language basics, and then touch on the 2 more advanced aspects of using VBA and Excel together.

Finally, it's worth noting that this basic overview misses a couple parts of VBA (TODO: like what) that might come up from time to time. They will be mentioned at the end of the chapter in passing, but this book is not a VBA reference. This book is designed to get you using VBA in sa professional setting with confidence. Knowing every nook and cranny of the language is not critical for that goal.
