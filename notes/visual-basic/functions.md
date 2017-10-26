# VBA Language Overview

# Functions

Functions, or "procedures" define a subset of application logic that will be executed when the function is invoked.

Functions are commonly used to perform actions or otherwise operate on objects or other variables.

Functions are like the "verb" to the object's "noun".

### Defining Functions

```vb
Private Sub MyFunction()
  ' do stuff here
End Sub
```

Function definitions begin with the statement `Private Sub`, followed on the same line by the name of the function (in this case `MyFunction()`), followed by one or more lines of indented code, and finally concluding with the statement `End Sub`.

Note the trailing parenthesis in the function's name. They not only visually indicate this statement is a function, but they also serve as a space to pass parameters (see below).

### Invoking Functions

The code inside a function won't execute until/unless invoked. Functions are generally invoked when the user triggers an event (like a button click event), or when the user "runs" the program.

### Function Parameters

TBA - Many functions are defined without need for parameters. But sometimes functions need certain other information in order to do their job. In some cases, functions have sufficient access to global scope variables, but other times they need local scope variables. In cases like these, we pass local scope information to the function by using "parameters":

```vb
' TBA
```
