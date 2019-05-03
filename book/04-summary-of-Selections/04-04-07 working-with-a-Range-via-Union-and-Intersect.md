### working with a `Range` via `Union` and `Intersect`

You can perform set operations on multiple Ranges using Union and Intersect. Like all set operations, they correspond to differnet sections of a Venn Diagram. The simpler example is using `Union` since it will always return a new valid Range if it was fed valid Ranges to start. It works by growing the Range into a new Range that includes all previous objects rerferenced.

Intersect is a different beast because it is possible for it to return `Nothing` if the givne Ranges do not actually intersect. This is actually a very useful property if you are trying ot confirm whether or not a given cell is within in another Rnage.

TODO: add a pciture of set operatiojns

Some common examples of where these functiojns come up:

- Intersect is used with Events and other usability tasks to determine if a givne or slected Cell is within a target Range
- Interect is very useful with Offset and Resize to grab a new Range that contains a subset of data of the orignal Range wihtout having to worry about creating a new Range that includes cells not previosuly indlucded. IN this sense, Intersect only alllows a Range to get smaller.
- Uniojn can be very helpful when building a larger group to change all of their properties at once. This is quite nice because Excel will "batch" the caluclations if oyu change the `Value` all at once. This sam etechnique can b used to build a Rnage to delete

TODO: add Uniojn-Dlete example

TODO: add Intersect exampe to remove headers

TODO: add Intersect technque for Events and Sleection changed
