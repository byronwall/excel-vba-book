# SO item 102
Good afternoon,

I was wondering if there is a way to count the number of blank spaces in a cell and return that value using a combination of vlookups, if else statements, and possible the len function? For example, the reports that are generated Generate values like this:

"TopLocation" " SubLocation" " SubSubLocation"

My goal is to be able to automate some of these reports, but I feel like I'm missing one or more pieces to the puzzle.

Thank you.

----

Assuming a blank only refers to a simple space, , you can get this with a quick formula that takes the difference of the original length and the length after removing spaces.

**Formula** in B1 with data in A1, copied down to end of data

```
=LEN(A1)-LEN(SUBSTITUTE(A1," ", ""))

```

**Picture**

![picture of results](https://i.stack.imgur.com/yhRFY.png)

**Data** in case you don't believe the image, where space is now period

```
a.b
a...c
..a..c

```
