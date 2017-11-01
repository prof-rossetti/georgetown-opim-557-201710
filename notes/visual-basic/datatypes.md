# VBA Language Overview

## Datatypes

### Strings

The "string" datatype is used to represent words or text. A strinng must begin with an opening quotation mark (`"`) and end with a closing quotation mark (`"`) (e.g. `"Hello World"`).

```vb
Dim MyMessage As String
MyMessage = "Hello World"
MsgBox(MyMessage)
```

The most popular string operation is "concatenation", which assembles multiple strings into a single string. The operator to perform string concatenation is an ampersand (`&`). When concatenating strings with other strings, or even with variables, make sure to include space characters in the proper places or else your strings will run together without a space. For example, **all the following approaches are equivalent**:

```vb
Dim MyMessage As String
MyMessage = "Hello" & " " & "World" ' notice the separate space character
MsgBox(MyMessage)
```

```vb
Dim MyMessage As String
MyMessage = "Hello " & "World" ' notice the trailing space after the word Hello
MsgBox(MyMessage)
```

```vb
Dim MyMessage As String
MyMessage = "Hello" & " World" ' notice the leading space before the word World
MsgBox(MyMessage)
```

```vb
Dim FirstString As String
Dim SecondString As String
MyMessage = FirstString & " " & SecondString ' notice the separate space character in-between the two variables. just because you use variables to represent strings does not change your need to include space characters
MsgBox(MyMessage)
```
