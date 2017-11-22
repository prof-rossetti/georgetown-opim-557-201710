# Project 2 - Stock Trading Recommendation System

You own and operate a financial planning business which helps customers make investment decisions.

Your objective is to build yourself a tool to automate the process of providing your clients with stock trading recommendations.

Specifically, the system should accept one or more stock symbols as information inputs, and should provide a recommendation as to whether or not the client should purchase the given stock(s).

## Prerequisites

  + [Project 1 - Retirement Savings Calculator](/projects/savings-calculator/project.md)
  + [Arrays](/notes/visual-basic/datatypes/arrays.md) and [Splitting Strings](/notes/visual-basic/datatypes/strings.md#string-splitting)
  + [Writing Sheets of Data](/notes/visual-basic/excel-objects.md#the-worksheet-object)
  + [Computer Networks](/notes/computer-networks/notes.md), [APIs](/notes/software/apis.md) and [Microsoft WinHTTP Services](/notes/visual-basic/references/win-http/notes.md)
  + Alpha Vantage API - [Registration](https://www.alphavantage.co/support/#api-key) and [Documentation](https://www.alphavantage.co/documentation/)
  + [Detecting Substrings](/notes/visual-basic/datatypes/strings.md#substring-detection)

## Learning Objectives

  + Design and build a tool to aid or automate a decision-making process.
  + Use VBA to capture and validate user inputs.
  + Use VBA to issue HTTP requests to retrieve CSV-formatted data from an API.
  + Use VBA to write CSV-formatted data to one or more spreadsheets.
  + Use VBA to perform programmatic calculations to arrive at a final system output.

## Information Requirements

### Information Inputs

The system should prompt the user to input one or more stock symbols (e.g. `"MSFT"`, `"AAPL"`, etc.).

The system may optionally prompt the user to specify additional inputs such as risk tolerance and other trading preferences, as desired and applicable.

### Information Outputs

The system should produce a recommendation as to whether or not the client should buy the stock, and optionally what quantity to purchase. The recommendation for each symbol can be binary (e.g. `"Buy"` or `"No Buy"`), qualitative (e.g. a `"Low"`, `"Medium"`, or `"High"` level of confidence), or quantitative (i.e. some numeric rating scale) in nature.

If the system includes any prices in its final recommendation, they should be formatted as USD with a dollar sign (`$`) and two decimal places.

## Interface Requirements

The system should capture inputs using your choice of input mechanism, whether it be cell value(s), input box(es), or some other means.

The system should use an ActiveX command button click or some other event to trigger the recommendation process.

The system should write historical stock prices to one or more worksheet(s). If the system processes only a single stock symbol at a time, the system may use a single sheet named something like "outputs" or "stock-prices". Whereas if the system processes multiple stock symbols at a time, for each stock symbol, the system should write historical trading data to a corresponding worksheet that is named after the stock symbol. If writing multiple sheets of data, the system should have a way of cleaning-up to prevent uncontrolled proliferation of new sheets. It is encouraged (especially for single-symbol solutions), but not necessary for price values on the output sheet(s) to be formatted as currency.

The system should display final recommendations using your choice of output mechanism, whether it be cell value(s), message box(es), or some other means.

## Validation Requirements

The system should first perform preliminary user input validations. For example, it should ensure stock symbols are `String` datatypes and less than around six characters long.

Also, when the system makes an HTTP request for that stock symbol's trading data, if the stock symbol is not found or there is an error message returned by the API server, the system should display a friendly error message like "Sorry, couldn't find any trading data for that stock symbol", and it should stop program execution to allow the user to try again.

## Calculation Requirements

You are free to develop your own custom recommendation algorithm. This is perhaps one of the most fun and creative parts of this project. :smiley:

One simple example algorithm would be (in pseudocode): If the stock's latest closing price is less than 20% above its 52-week low, "Buy", else "Don't Buy".










## Submission Instructions

Submit a single macro-enabled excel file to [Blackboard](https://campus.georgetown.edu/webapps/assignment/uploadAssignment?content_id=_4454669_1&course_id=_745457_1&assign_group_id=&mode=cpview). The file should be named **project-2-`netid`.xlsm**, where `netid` represents your own university-issued Net Id (e.g. **project-2-abc123.xlsm**).

## Evaluation Methodology

Full credit for a properly-named file which accepts one or more stock-symbol user inputs, validates inputs, issues corresponding HTTP requests to the AlphaVantage API, handles response errors as appropriate, writes response data to one or more worksheets, and provides final purchase recommendation(s).

Else partial credit to highlight areas of improvement.

Note: The professor reserves the right to award extra credit in recognition of particularly-effective user experiences.

### Tentative Rubric

A tentative grading rubric is as follows:

top-level requirement | tentative weight
--- | ---
File Naming | around 4%
Information Requirements | around 24%
Interface Requirements | around 24%
Validation Requirements | around 24%
Calculation Requirements | around 24%
