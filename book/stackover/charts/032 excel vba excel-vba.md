# SO item 032
I'm not sure this is possible but thought this was the best place to ask.

Is it posible to get the position of a series value on a graph in excel?

For example, if I have a line graph in excel that has time along the x axis, is it possible to (using VBA) get the position of a specific point on that axis.

What I am trying to do is have a vertical line that is can be positioned based on a date entered by the user.

like this

![enter image description here](https://i.stack.imgur.com/n1G6U.jpg)

Where the green line could be positioned by entering in a date (rather than just being manually moved) (or also it could be set to automatically move to the current date etc).

I was then thinking that if the position is on the graph is queryable, then I can just access the line object and move it to any position I wanted through VBA.

Any Ideas? or is this just not possible?

----

The "cleanest" way to do this is to add the line to the chart as a new series. In that way, Excel handles all of the positioning and your work is simplified. To get a vertical line on a chart, there are a number of options. I prefer this route:

1.  Create a small 2x2 area with two dates and two values
2.  Add in the date or x-axis value you want the line at (`E3` in image). You can use `=TODAY()` here or some manually entered value.
3.  Set the second x-axis value equal to the first
4.  Use `MAX` and `MIN` on the data to get the values for each date. You can also use 0 and 1 and a secondary axis, but I think `MAX/MIN` is easier.
5.  Add the data to the chart and format as a marker with straight line.

**Formulas**

*   `E3`: `=TODAY()`
*   `E4`: `=E3`
*   `F3`: `=MIN(C3:C27)`
*   `F4`: `=MAX(C3:C27)`

**Result and chart data series for vertical line**

![results and chart](https://i.stack.imgur.com/ELSZS.png)
