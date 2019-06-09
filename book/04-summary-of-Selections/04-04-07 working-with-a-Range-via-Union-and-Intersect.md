### working with a `Range` via `Union` and `Intersect`

You can perform set operations on multiple Ranges using Union and Intersect. Like all set operations, they correspond to different sections of a Venn Diagram. The simpler example is using `Union` since it will always return a new valid Range if it was fed valid Ranges to start. It works by growing the Range into a new Range that includes all previous objects referenced.

Intersect is a different beast because it is possible for it to return `Nothing` if the given Ranges do not actually intersect. This is actually a very useful property if you are trying to confirm whether or not a given cell is within in another Range.

TODO: add a picture of set operations

Some common examples of where these functions come up:

- Intersect is used with Events and other usability tasks to determine if a given or selected Cell is within a target Range
- Interacted is very useful with Offset and Resize to grab a new Range that contains a subset of data of the original Range without having to worry about creating a new Range that includes cells not previously indlucded. IN this sense, Intersect only allows a Range to get smaller.
- Union can be very helpful when building a larger group to change all of their properties at once. This is quite nice because Excel will "batch" the calculations if oyu change the `Value` all at once. This sam etechnique can b used to build a Range to delete

TODO: add Union-Dlete example

TODO: add Intersect exampe to remove headers

TODO: add Intersect technque for Events and Sleection changed
