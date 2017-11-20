# Project 2 - Stock Trading Recommendation System

You own and operate a financial planning business which helps customers make investment decisions.

Your objective is to build yourself a tool to automate the process of providing your clients with stock trading recommendations.

Specifically, the system should accept one or more stock symbols as information inputs, and should provide a recommendation as to whether or not the client should purchase the given stock(s).

## Prerequisites

  + [Project 1 - Retirement Savings Calculator](/projects/savings-calculator/project.md)
  + [Arrays](/notes/visual-basic/datatypes/arrays.md) and [Splitting Strings](/notes/visual-basic/datatypes/strings.md#string-splitting)
  + [Writing Sheets of Data](/notes/visual-basic/excel-objects.md#the-worksheet-object)
  + [APIs](/notes/software/apis.md) and [Web Requests](/notes/visual-basic/web-requests.md)
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

## Interface Requirements

The system should capture inputs via cell values, input boxes, or ActiveX controls.

The system should use an ActiveX command button click or some other event to trigger the recommendation process.

For each stock symbol input by the user, the system should write historical trading data to a corresponding worksheet that is named after the stock symbol.

The system should provide final recommendations via message box, cell values, or some other means.

## Validation Requirements

The system should first validate user inputs (for example, ensuring stock symbols are `String` datatypes and less than around six characters long).

Also, when the system makes an HTTP request for that stock symbol's trading data, if the stock symbol is not found, the system should display a friendly error message like "Sorry, couldn't find any trading data for that stock symbol".

## Calculation Requirements

You are free to develop your own custom decision-making algorithm. This is perhaps one of the most fun and creative parts of this project.

One simple example algorithm would be (in pseudocode): If the stock's latest closing price is less than 20% above its 52-week low, "Buy", else "Don't Buy".










## Submission Instructions

Submit a single macro-enabled excel file to [Blackboard](https://campus.georgetown.edu/webapps/assignment/uploadAssignment?content_id=_4454669_1&course_id=_745457_1&assign_group_id=&mode=cpview). The file should be named **project-2-`NETID`.xlsm**, where `NETID` represents your own university-issued Net Id (e.g. **project-2-abc123.xlsm**).

## Evaluation Methodology

Full credit for a system which accepts one or more stock-symbol user inputs, issues corresponding HTTP requests to the AlphaVantage API, writes the resulting data to one or more worksheets, and provides final purchase recommendation(s).

Else partial credit to highlight areas of improvement.

Note: The professor reserves the right to award extra credit in recognition of particularly-effective user experiences.
