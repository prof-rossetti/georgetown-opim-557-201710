# Project 1 - Savings Model

## Prerequisites

  + ["All the Controls" Exercise](/exercises/all-the-controls/exercise.md) (and all of its prerequisites)
  + Control Flow - [Loops](/notes/visual-basic/loops.md)
  + Data Quality - [Validations](notes/visual-basic/datatypes.md)

## Learning Objectives

  + Design and build a tool to aid a decision-making process.
  + Use ActiveX controls and VBA to capture and validate user inputs.
  + Use VBA to perform programmatic calculations to arrive at a final system output.

## Requirements

You run a financial planning business which helps customers plan for their retirement.

Your objective is to build yourself a system to automate the common calculations required to provide your clients with retirement savings advice.

Specifically, the system should accept a number of information inputs representing the client's savings goals, and should produce an information output representing the amount of money the client can expect to have saved upon reaching retirement age.

### Information Inputs

Your system should accept the following user inputs:

  + The client's current age.
  + The client's desired retirement age.
  + The client's current amount of savings (assume the client does not have any debt).
  + The client's current salary (assume the client earns income from a salary and no other investment vehicles besides his/her savings).
  + A projected annual growth rate for the client's salary.
  + A projected annual grown rate for the client's savings.

The table below provides a framework for you to translate these information inputs into variables. The min and max variable values are just reasonable suggestions, and in some cases when indicated as being "flexible", can be modified based on your own preference.

info input | suggested variable name | variable datatype | default value | min allowable value | max allowable value
--- | ---  | ---  | ---  | ---  | ---
Current Age | `Age` | `Integer` | `30` | `18` (you don't give advice to minors) | `60` (flexible)
Desired Retirement Age | `RetirementAge` | `Integer` | `65` | `35` (flexible) | `80` (flexible)
Savings Balance | `SavingsBalance` | `Double` | `10000.00` | `0.00` | `60000.00` (flexible)
Current Salary | `CurrentSalary` | `Double` | `80000.00` | `40000.00` (flexible) | `250000.00` (flexible)
Annual Salary Growth Rate | `SalaryGrowthRate` | `Double` | `0.03` | `0.00` | `0.15` (flexible)
Annual Savings Growth Rate | `SavingsGrowthRate` | `Double` | `0.01` (flexible) | `0.005` (flexible) | `0.10` (flexible)

### Information Outputs

Your system should produce the following outputs:

  + The amount of savings projected at the client's specified retirement age.



## Instructions

Design an interface, capture user inputs, validate user inputs, perform calculations, and produce the desired output.

### Interface Design

TBA - hints and specific instructions forthcoming

### Capturing User Inputs

TBA - hints and specific instructions forthcoming

### Validating User Inputs

TBA - hints and specific instructions forthcoming

### Performing Calculations

TBA - hints and specific instructions forthcoming

### Producing Output

TBA - hints and specific instructions forthcoming








## Submission Instructions

Submit a single macro-enabled excel file to [Blackboard](https://campus.georgetown.edu/webapps/assignment/uploadAssignment?content_id=_4454661_1&course_id=_745457_1&assign_group_id=&mode=cpview). The file should be named after your university-issued Net Id (e.g. `abc123.xlm`)

## Evaluation Methodology

Full credit for a properly-named file containing a user-friendly interface which properly validates all user inputs and performs the proper calculations to produce the correct output value.

Else partial credit to highlight areas of improvement.
