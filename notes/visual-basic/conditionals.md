# VBA Language Overview

## Control Flow

### Conditionals

#### `If ... End If` Statements

Reference: [documentation](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/ifthenelse-statement).

Examples:

```vb
If CheckBox1.Value = True Then
  MsgBox("The check box has been selected")
End If
```

```vb
If CheckBox1.Value = True Then
  MsgBox("The check box is selected")
Else
  MsgBox("The check box is not selected")
End If
```
