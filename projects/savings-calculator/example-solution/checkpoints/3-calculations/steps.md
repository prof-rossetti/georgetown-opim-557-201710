# Project 1 - Retirement Savings Calculator

## Checkpoint 3 (Calculations) Walk-through

### Setup

Download the ["calculation-less" example solution](/projects/savings-calculator/example-solution/example-solution-calculationless.xlsm).

Reference also copies of the solution's [VBA files](/projects/savings-calculator/example-solution/vba-files).

For any given interface example, find the section titled `"CALCULATE OUTPUTS"`, specifically the `"perform calculations here"` placeholder.

For each step below, replace the `"perform calculations here"` placeholder with the code contained in that section. Then test the behavior of the program. When you are satisfied, replace the previous step's code with the next step's code and repeat the process.

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
