### finding a `Range` while iterating through a `Range`

One technique for working with Ranges is to start with one Range, iterate though it, can build a new Range based on some criteria. Alternatively, you may just act immdiatel on the Range as you are iterating through it. This approach is dead simple and is used in abundance htoruhgout good workflows. As long as there is some meaningful logic which can be applied to identify whether or not a subset of a Range is interesting, you can use this technique. SOme common logical stepst aht are chcked:

- Check the `Value` of the cell
- Check if the cell has some property (e.g. `HasFormula`, `HasArray`, etc.)
- Check the `Style` of the cell

The idea is simple: check some property while iterating and act on it. THis is obvious once you have been programming for a bit, but sometimes you just need to be told htat this is an acceptable way of doing things. You do not always need to use `Find` to search for a cell that contains some value. You can always just iterate all the cells and see if a cell matches that value (or contains it with `InStr`).

TODO: find some code related to this?
