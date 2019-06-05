### Intersect

The first is the `Intersect` technique to determine if a Range that was affected by an event was a Range of interest. With this approach, you define a Range which includes your "interesting" cells. You then do a `If Not Intersect(rngEvent, rngTarget) Is Nothign` to see if the interesection of the callback Range and the desired Range overalp. If they overlap, yhen you typically execute some code. This allows you to quickly filter out Ranges which have hcanged but are not reelvenat ot ahtever code you need to run.

TODO: add a code sample here
