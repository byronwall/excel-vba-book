## creating charts from scratch

The previous section discussed how to work with exisitng Charts.  This section will focus on how to create those Charts from scratch if you are coming into a blank Worksheet or if you simply need to add a chart to existnig data.  At the start, it's worth mentioning that creating Cahrts from scrathc falls into one of two categories:

* Library/helper type code where you want to quickly create a Chart in a common way.  This type of code works best in an addin and typicalyl provides funcitonaltiy that you wish Excel had from teh start
* One-off code for a specific application.  This invovles creating a Chart with some sort of odd manipualation or formatting or otehr detail where automation saves time.

The two types of category will end up with code that looks similar, but the goals of the former cateogry will be slightly different than teh latter.  Typically when making code for a one-off applicatio, you can make more assumptions baout how the data is structued and what sorts of actions need ot be taken.  When working with helper code, you will spend more time asking for user input, and handling the different cases that might come up.

Another key point to make is that the type of work that is being done in a chart can vary as well.  The splitting line here is whether the Chart creation is data heavy or formatting heavy (or possibly both).  For a data heavy Chart, you will spend a lot of time collecting Ranges, creating Series, and possibly maniupalting individual Points.  For a formatting heavy chart, you will spend a lot of time iteratng through the Series to apply formatting, label the Axes, set the number formats, and generally modify the Excel defaults.  Both of these tasks are very time intensive if you are doing them without VBA, so both lend themselves to being automated if possible.

Excel provides two means of creating a Chart depending on how you want to handle things.  Those two commands are:

* ChartObjects.Add
* TODO: waht is the other method

I always prefer to use ChartObjects.Add because of it consistent applicaiton.  The other approach tends to put you at the mercy of how Excel interprets your data and its layout.

TODO: add more detail here

The general process for creating a chart looks like this:

* Create a new ChartObject via ChartObjects.Add - store that reference
    * If you know where you want the Chart to go, you can use that information here
* Access the Chart of that object
* Change the properties of the Chart that you know -- namely ChartType
* Access the SeriesCollection of the Chart and call NewSeries for each Series you want - store a reference to that Series
    * This is typically done inside a loop that is iterating through Ranges in some way
    * If you need to apply Series specific formatting, do that here
* Access the Axes colleciton and modify any specific parts of the Axes that you want
    * This may show up in the loop above if you want the Axis to draw information from teh Series (maybe set the max to the max of the data?)

At this point, you will have a Chart with the Series you want along with the major formatting taken care of.  Even better, this general framework lends itself nicely to adding new commands where needed.  If you need to go after some of the finer details of the Chart, you can add those comamnds where the objects are being reference, or at teh end of the code.  The main thing to consider is whether you need to work insnide loops (per Series) or if you can process the extra stuff at the end.

The other upside of this approach is that you can quickly wrap all of this code with another loop to create multiple Charts.  You can then wrap that code with another loop to do it on multiple Worksheets, etc.  When you write code that can cleanly live inside a loop, you make it easy to use the code elsewhere.

One other aspect of Charts that is somewaht unique is that you can typically reuse a lot of the code by creating new Subs.  These can be called from the inside of a loop to create a chain of commands to process a Chart.  This approach is highly effective if you work in an environmetn where the same or similar things need to be done.  For example: you have a monthyl report to create each month for multiple departments.  Standardizing as much of that work into modules makes it easy to apply teh code in multiple spots with minor changes.  This is relevant to Charts becuase most of the work of Charts is changingthe values of specific propreties.  There is typicallly far less logic that is unique to an applicaiton (like trying to build a Range based on teh layout of data).

Once you have this general frameowkr mastered, you can quickly use it to make more charts.

TODO: add some examples of creating Charts
