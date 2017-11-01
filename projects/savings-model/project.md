# Project 1 - Savings Model

You run a financial planning business which helps customers plan for their retirement.

Your objective is to build yourself a system to automate the common calculations required to provide your clients with retirement savings advice.

Specifically, the system should


## Prerequisites

  + ["All the Controls" Exercise](/exercises/all-the-controls/exercise.md) (and all of its prerequisites)
  + Control Flow: [Loops](/notes/visual-basic/loops.md)
  + Data Quality: [Validations](notes/visual-basic/datatypes.md)

## Learning Objectives

  + Demonstrate ability to use ActiveX Controls
  + B
  + C

## Requirements

### Information Inputs

Your system should accept the following inputs:

  + The client's current age.
  + The client's desired retirement age.
  + The client's current amount of savings (assume the client does not have any debt).
  + The client's current salary (assume the client earns income from a salary and no other investment vehicles besides his/her savings).
  + A projected annual growth rate for the client's salary.
  + A projected annual grown rate for the client's savings.

### Information Outputs

Your system should ____ the following outputs:

  + The amount of savings projected at the client's specified retirement age.


## Instructions

This section contains optional instructions to help you get started in planning your system development process.


TBA

The table below provides a framework for you to translate these information inputs into variables. The min and max variable values are just reasonable suggestions, and in some cases when indicated as being "flexible", can be modified based on your own preference.

info input | suggested variable name | variable datatype | default value | min allowable value | max allowable value
--- | ---  | ---  | ---  | ---  | ---
Current Age | `Age` | `Integer` | `30` | `18` (you don't give advice to minors) | `60` (flexible)
Desired Retirement Age | `RetirementAge` | `Integer` | `65` | `35` (flexible) | `80` (flexible)
Savings Balance | `SavingsBalance` | `Double` | `10000.00` | `0.00` | `60000.00` (flexible)
Current Salary | `CurrentSalary` | `Double` | `80000.00` | `40000.00` (flexible) | `250000.00` (flexible)
Annual Salary Growth Rate | `SalaryGrowthRate` | `Double` | `0.03` | `0.00` | `0.15` (flexible)
Annual Savings Growth Rate | `SavingsGrowthRate` | `Double` | `0.01` (flexible) | `0.005` (flexible) | `0.10` (flexible)

## Submission Instructions

TBA

## Evaluation Methodology

TBA
