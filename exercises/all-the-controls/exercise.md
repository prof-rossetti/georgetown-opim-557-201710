# "All the Controls" Exercise

## Prerequisites

  + ["Self-aware Button" Exercise](/exercises/self-aware-button/exercise.md)
  + [Excel Objects Overview](/notes/visual-basic/excel-objects.md)
  + [ActiveX Controls Overview](/notes/visual-basic/activex-controls.md)
  + [`If ... End If` statements](/notes/visual-basic/conditionals.md)

## Learning Objectives

  + Initialize and configure the most popular ActiveX Controls.
  + Programmatically access a control's properties.
  + Respond to events in a control's event lifecycle.

## Challenges

### More `CommandButton` Challenges:

  1. Make a `CommandButton` that when clicked alerts the user of the value of some specified cell.
  1. Make a `CommandButton` that when clicked changes the value of some specified cell.
  1. Make a `CommandButton` that when clicked changes the value of some specified cell, and alerts the user of the cell's old and new values, respectively. Hint: store the old value in a variable before changing it.

![a screenshot of a message box which displays a cell's value (456)](/notes/visual-basic/activex-controls/command-button/command-buttons-get.png)

![a screenshot of a message box which has overwritten a cell's value from 456 to 123.](/notes/visual-basic/activex-controls/command-button/command-buttons-set.png)

### `ToggleButton` Challenges

  1. Make a `ToggleButton` that uses a linked cell to display whether or not the button is currently pressed (i.e. `True` or `False`).
  1. Make a `ToggleButton` that when pressed alerts the user.
  1. Make a `ToggleButton` that when pressed alerts the user of whether or not the button is currently pressed (i.e. `True` or `False`).
  1. Make a `ToggleButton` that when pressed alerts the user of whether it has been "pressed" or "unpressed". Hint: use an `IF ... End If` statement.

![a screenshot of a message box displaying the button has been toggled "on". it also uses a linked cell to display its boolean value.](/notes/visual-basic/activex-controls/toggle-button/toggle-button-clicked-on.png)

### `ComboBox` Challenges

  1. Make a `ComboBox` that allows the user to select an option from a provided list.
  1. Make a `ComboBox` that allows the user to select an option from a provided list and uses a linked cell to display the currently selected item.
  1. Make a `ComboBox` that allows the user to select an option from a provided list and alerts the user when an item is selected.
  1. Make a `ComboBox` that allows the user to select an option from a provided list and alerts the user which item was selected.

![a screenshot of a user selecting an option from a drop-down menu.](/notes/visual-basic/activex-controls/combo-box/combo-box-1.png)

![a screenshot of a message box displaying the name of an item that has been selected from a drop-down menu. also it displays the selected value in a linked cell.](/notes/visual-basic/activex-controls/combo-box/combo-box-2.png)

### `ListBox` Challenges

Repeat the `ComboBox` challenges (see above), but use a `ListBox` control instead.

![a screenshot of a list box control which displays the currently selected item name in a linked cell](/notes/visual-basic/activex-controls/list-box/list-box.png)

### `SpinButton` Challenges

  1. Make a `SpinButton` that allows the user to increment or decrement an integer value between some specified acceptable range of values.
  1. Make a `SpinButton` that allows the user to increment or decrement an integer value between some specified acceptable range of values and uses a linked cell to display the currently selected value.
  1. Make a `SpinButton` that allows the user to increment or decrement an integer value between some specified acceptable range of values and alerts the user when a value is incremented or decremented.
  1. Make a `SpinButton` that allows the user to increment or decrement an integer value between some specified acceptable range of values and alerts the user which value was selected.

![a screenshot of a message box which displays the current integer value of a spin button control. also it displays its value in a linked cell.](/notes/visual-basic/activex-controls/spin-button/spin-button-increment.png)

### `ScrollBar` Challenges

Repeat the `SpinButton` challenges (see above), but use a `ScrollBar` control instead.

![a screenshot of a message box which displays the current integer value of a scroll bar control. also it displays its value in a linked cell.](/notes/visual-basic/activex-controls/scroll-bar/scroll-bar-scrolled.png)

### `OptionButton` Challenges

  1. Make four `OptionButton` controls belonging to the same group, each having its own name and caption.
  1. Make four `OptionButton` controls belonging to the same group, each having its own name and caption, and each using a linked cell to indicate whether or not it has been selected (i.e. `True` or `False`). Clarification: use a different cell for each control.
  1. Make four `OptionButton` controls belonging to the same group, each with its own name and caption, such that when any one of the options is selected it alerts the user which option has been selected.
  1. Make four `OptionButton` controls belonging to the same group, each with its own name and caption, such that when any one of the options is selected it writes its caption to a specified cell. Clarification: use the same cell for all controls.

![a screenshot of four vertically-aligned option buttons, one of which is selected. it also shows a message box alerting the user of which option has been selected. it also uses four different linked cells to indicate the boolean values of each option. it also displays in a specified cell the caption of the selected option.](/notes/visual-basic/activex-controls/option-button/option-button-2.png)

### `CheckBox` Challenges

Repeat the `OptionButton` challenges (see above), but use `CheckBox` controls instead. For challenge #4, instead of writing the caption of a single selected option to a specified cell, write a concatenated list of all selected options. Hint: use an `If ... End If` statement.

![a screenshot of four vertically-aligned check boxes, two of which are selected. it also shows a message box alerting the user of which options have been selected. it also uses four different linked cells to indicate the boolean values of each option. it also displays in a specified cell the captions of both selected options.](/notes/visual-basic/activex-controls/check-box/check-box-2.png)
