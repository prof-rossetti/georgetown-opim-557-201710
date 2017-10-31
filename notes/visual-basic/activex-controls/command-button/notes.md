# MS Excel ActiveX Controls

## The `CommandButton` Control

A button to be clicked.

Reference: [documentation](https://msdn.microsoft.com/en-us/VBA/Language-Reference-VBA/articles/commandbutton-control).

### Initialization

"Developer" > "Insert" > "ActiveX Controls" > "Command Button"

![a screenshot of an excel worksheet with two buttons which read "Get cell value" and "Set cell value", respectively.](command-button.png)

### Properties

name | description
--- | ---
`Caption` | Human-friendly text to instruct the user.

### Events

name | description
--- | ---
`Click` (default) | Triggers when the button is clicked.
