'
' A SOLUTION FOR ALL-CELLS INTERFACE
' ... Author: Prof. Rossetti <prof.mj.rossetti@gmail.com>.
' ... License: Students, feel free but not obligated to use this code in your project as long as you retain this attribution section. If you wrote something like this on your own, no need to attribute. If this code inspired you to write your own code, please still consider providing an attribution link to this file's GitHub URL.
'

Private Sub CommandButton1_Click()
    Dim Age
    Dim RetirementAge
    Dim SavingsBalance
    Dim AnnualContribution
    Dim AnnualInterestRate

    '
    ' CAPTURE USER INPUTS (VIA CELL VALUES)
    '

    Age = Range("E9").Value
    RetirementAge = Range("E11").Value
    SavingsBalance = Range("E13").Value
    AnnualContribution = Range("E15").Value
    AnnualInterestRate = Range("E17").Value

    '
    ' VALIDATE USER INPUTS
    '

    If IsValidAge(Age) = False Then Exit Sub
    If IsValidAge(RetirementAge) = False Then Exit Sub
    If AgesValid(Age, RetirementAge) = False Then Exit Sub
    If IsValidUSD(SavingsBalance) = False Then Exit Sub
    If IsValidUSD(AnnualContribution) = False Then Exit Sub
    If IsValidPct(AnnualInterestRate) = False Then Exit Sub

    '
    ' DISPLAY USER INPUTS
    '

    Call LogUserInputs(Age, RetirementAge, SavingsBalance, AnnualContribution, AnnualInterestRate)

    '
    ' CALCULATE OUTPUTS
    '

    Dim TotalContribution As Double
    Dim TotalInterest As Double

    ' ... perform calculations here (see checkpoint steps)

    '
    ' DISPLAY FINAL OUTPUTS
    '

    Call LogFinalOutputs(SavingsBalance, TotalContribution, TotalInterest)
End Sub
