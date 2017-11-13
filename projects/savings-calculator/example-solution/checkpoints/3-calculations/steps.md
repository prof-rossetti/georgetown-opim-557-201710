# Project 1 - Retirement Savings Calculator

## Checkpoint 3 (Calculations) Walk-through

### Setup

```vb
'
' CALCULATE OUTPUTS
' ... Author: Prof. Rossetti <prof.mj.rossetti@gmail.com>.
' ... License: Students, feel free but not obligated to use this code in your project as long as you retain this attribution section. If you wrote something like this on your own, no need to attribute. If this code inspired you to write your own code, please still consider providing an attribution link to this file's GitHub URL.
'

Dim TotalContribution As Double ' need to display this (not relevant until Step 4)
Dim TotalInterest As Double ' need to display this (not relevant until Step 4)

' ... perform calculations here (see steps, below)
```

Keep the setup code above, and swap in each of the following steps in succession.

### Step 1

Calculate savings balance for end of first year:

```vb
SavingsBalance = SavingsBalance * (1 + AnnualInterestRate)
SavingsBalance = SavingsBalance + AnnualContribution
```

### Step 2

Loop through each year between current age and retirement age:

```vb
Do While (Age <= RetirementAge)
    MsgBox ("Age: " & Age)

    Age = Age + 1 ' increment the age to avoid infinite loop!
Loop
```

### Step 3

Calculate final savings balance:

```vb
Do While (Age <= RetirementAge)
    SavingsBalance = SavingsBalance * (1 + AnnualInterestRate)
    SavingsBalance = SavingsBalance + AnnualContribution

    MsgBox ("Age: " & Age & vbNewLine & "Balance: " & FormatUSD(SavingsBalance) & ".")

    Age = Age + 1 ' increment the age to avoid infinite loop!
Loop
```

### Step 4

Calculate all final outputs:

```vb
TotalContribution = SavingsBalance ' count initial savings balance toward total contribution

Do While (Age <= RetirementAge)
    AnnualInterest = SavingsBalance * AnnualInterestRate
    SavingsBalance = SavingsBalance + AnnualInterest + AnnualContribution

    TotalContribution = TotalContribution + AnnualContribution ' keep track of total contribution
    TotalInterest = TotalInterest + AnnualInterest ' keep track of total accrued interest

    Age = Age + 1 ' increment the age to avoid infinite loop!
Loop
```
