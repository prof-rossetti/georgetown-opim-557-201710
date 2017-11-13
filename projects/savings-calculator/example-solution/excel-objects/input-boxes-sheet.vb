'
' A SOLUTION FOR ALL-INPUT-BOXES INTERFACE
' ... Author: Prof. Rossetti <prof.mj.rossetti@gmail.com>.
' ... License: Students, feel free but not obligated to use this module in your project as long as you retain this attribution section. If you wrote something like this on your own, no need to attribute. If this code inspired you to write your own code, please still consider providing an attribution link to this file's GitHub URL.
'

Private Sub CommandButton1_Click()
    Dim Age
    Dim RetirementAge
    Dim SavingsBalance
    Dim AnnualContribution
    Dim AnnualInterestRate

    '
    ' CAPTURE USER INPUTS (VIA NUMERIC-TYPE INPUT BOXES)
    ' ... AND VALIDATE INPUTS IMMEDIATELY AFTER EACH IS CAPTURED
    ' ... (FOR BETTER USER EXPERIENCE)
    '

    Age = Application.InputBox(prompt:="Please specify your current age (e.g. 60): ", Type:=1)
    If IsValidAge(Age) = False Then Exit Sub

    RetirementAge = Application.InputBox(prompt:="Please specify your desired retirement age (e.g. 65): ", Type:=1)
    If IsValidAge(RetirementAge) = False Then Exit Sub
    If AgesValid(Age, RetirementAge) = False Then Exit Sub

    SavingsBalance = Application.InputBox(prompt:="Please specify your current savings balance (e.g. 50000.00): ", Type:=1)
    If IsValidUSD(SavingsBalance) = False Then Exit Sub

    AnnualContribution = Application.InputBox(prompt:="Please specify your predicted annual contribution (e.g. 18000.00): ", Type:=1)
    If IsValidUSD(AnnualContribution) = False Then Exit Sub

    AnnualInterestRate = Application.InputBox(prompt:="Please specify your predicted annual interest rate (e.g. 0.05): ", Type:=1)
    If IsValidPct(AnnualInterestRate) = False Then Exit Sub

    '
    ' DISPLAY USER INPUTS
    '

    Call LogUserInputs(Age, RetirementAge, SavingsBalance, AnnualContribution, AnnualInterestRate)

    '
    ' CALCULATE OUTPUTS
    '

    ' ... perform calculations here (see checkpoint steps)

    '
    ' DISPLAY FINAL OUTPUTS
    '

    Call LogFinalOutputs(SavingsBalance, TotalContribution, TotalInterest)
End Sub
