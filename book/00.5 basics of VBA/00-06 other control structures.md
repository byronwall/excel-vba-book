## other control structures

### With command

THe `With` command allows you to place a given variable within "scope" and avoid repeatedly typing that variable's name for each required call.  The `With` command exists solely to reduce the nujmber of times that a givne object or variable name is typed.  You are never required to ues a With command to accomplish a goal, but it can be helpful to clarify or avoid having too long of a code block.  Having said that, a With block can be incredibly confusing to read especially when mixed with the always in scope function calls like `Range` or `Cells`.  It is incredibly easy to avoid typing the required `.` to start a new line and accidentally refer to the globally scope object instead of your With scoped object.  For this reason, I very rarely use the With command. When I do use it, I will tpyically only use it when I am workign with a nested object that might be several levles deep.  Having said that, I mostly avoid the With block by creating a variable which holds the object in question adn using that instead.  I have found that parsing a With block later can quickly become a confusing mess becuase of the difficulty of spotting the `.` which is critical.

If you read through some of the most common questons on teh interent about "why my VBA no work?" you will quickly find issues with With blocks accidentally calling a globally scoped command.  I have never asked those questions on the internet, but I have definitely been bittne by teh same errors where a `.` is missed and the commang goes bonkers.  It happens but is easily avoided by not using `With`.

TODO: add some content about With

TODO: add some content about Goto and labels

TODO: add some content about Error handling
