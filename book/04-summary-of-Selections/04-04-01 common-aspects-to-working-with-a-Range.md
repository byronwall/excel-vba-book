### common aspects to working with a `Range`

There are several common aspects of working with `Ranges`. The most important thing is to remember the difference between using the `Range` as a reference or as a `Value`. The problem comes because VBA will work really hard to allow your code to execute regardless of whether the Value/reference part is done correctly.

The difference is best explained with an example. In this example, you can see that when the reference is stored, you must use the `Set` command. If you want the `Value` of a `Range`, you can either use `Value` or rely on VBA calling it implicitly otherwise. If you attempt to assign the `Value` of a `Range` to a `Range` object, you will get an error. If you attempt to assign the `Value` of a `Range` to a `Variant` variable, it will work, but the variable will only hold the `Value`. That is, you cannot make further calls from the `Range` object model. This should highlight the importance of declaring variables with the tightest scope on the variable type. If everything is a `Variant`, VBA will let you get away with a lot; sometimes that flexibility will bite you.

TODO: add an example here
