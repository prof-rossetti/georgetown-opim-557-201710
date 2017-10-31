# MS Excel ActiveX Controls

## The `CheckBox` Control

A checkable box belonging to a specified group from which zero or more may be selected at any given time.

Reference: [documentation](https://msdn.microsoft.com/en-us/VBA/Language-Reference-VBA/articles/checkbox-control).

### Initialization

For each box: "Developer" > "Insert" > "ActiveX Controls" > "Check Box".

![a screenshot depicting two of four checked boxes](check-box.png)

### Properties

name | description
--- | ---
`Caption` | a human-friendly name for the selectable option.
`GroupName` | Associates the control with a logical grouping of one or more controls (default: "Sheet1").
`Value` | The name of the currently-selected list item.
`LinkedCell` | The address of a specified cell which is bidirectionally associated with control's value.

### Events

name | description
--- | ---
`Click` (default) | Triggers when an option is selected from the from the list.
`Change` | Triggers when an the control's value is changed.
