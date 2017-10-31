# "Self-aware Button" Exercise - Counting Button Clicks (Solution)

This solution is a lesson in variable scope.

## Iterative Development Approach

Focusing on user experience, how would the message look?

```vb
Private Sub CommandButton1_Click()
  MsgBox("Hello User! You have clicked me 1 time(s).")
End Sub
```

Abstract-away the click-count concept into a variable, because we expect it to change. It's ok if we're not changing it yet:

```vb
Private Sub CommandButton1_Click()
  Dim ClickCount As Integer
  ClickCount = 1
  MsgBox("Hello User! You have clicked me " & ClickCount & " time(s).")
End Sub
```

The `ClickCount` variable still doesn't update/change. Not knowing anything about variable scope, we might try the following approach to get closer to the desired behavior:

```vb
Private Sub CommandButton1_Click()
  Dim ClickCount As Integer ' remember: the default value for Integer variables is zero (0)
  ClickCount = ClickCount + 1
  MsgBox("Hello User! You have clicked me " & ClickCount & " time(s).")
End Sub
```

... but the click count still doesn't increment as desired. This is because every time the `CommandButton1_Click()` function gets executed, it re-declares the `ClickCount` variable and resets its value.

At the moment, the `ClickCount` variable is said to be **local-scope**. That is, it is defined and assigned only inside the `CommandButton1_Click()` function. Its value will not carry through to other functions or subsequent invocations of that function. Its value is erased from the program's memory after its function finishes execution.

Instead, we need to declare a **global-scope** variable whose value will remain in the program's memory even after the function has finished execution:

```vb
Dim ClickCount As Integer ' moving the variable declaration outside of the function into the "global scope" ensures its value can be accessed across various functions

Private Sub CommandButton1_Click()
  ClickCount = ClickCount + 1
  MsgBox("Hello User! You have clicked me " & ClickCount & " time(s).")
End Sub
```

Nice Job!
