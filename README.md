
# SpreadJSON

Have you ever worked on your `Google Sheets/Apps Script` and used functions like `getRange`, `setValue`, `setValues`, etc.?

This JavaScript script is designed to enable you to perform those actions directly on `Node.js`.

`SpreadJSON` utilizes the popular `SheetJS` Node.js library to store your sheet in JSON format and provides familiar methods.

Thus, this script simply allows you to store your sheet in JSON format and read or write to the sheet directly on Node.js.

(It can be highly beneficial if you have an Apps Script project and want to migrate to Node.js for faster execution and flexibility.)

[Caution] This project is still incomplete. There are many features to add to make it more general and easy to use.

## Prerequisites

This project requires NodeJS.
[Node](http://nodejs.org/) are really easy to install.
To make sure you have them available on your machine,
try running the following command.

```sh
$ node -v
v20.10.0
```
> This script use 3 dependencies which is `SheetJS` `fs` and `rwlock`.<br>
> Rwlock was used to implement a basic solution to \`race condition\` problem.
> 
And you have to install above 3 libraries for SpreadJSON to work

## Table of contents

- [SpreadJSON](#SpreadJSON)
  - [Prerequisites](#prerequisites)
  - [Table of contents](#table-of-contents)
  - [Getting Started](#getting-started)
  - [Usage](#usage)
    - [Example](#example)

## Getting Started

## Usage

Simply put SpreadJSON.js file on your project and include to your workspace.

[Caution]
<br>This script can only work for SheetJS's workbook that need to be in the format of JSON.
<br>You can use `XLSX.utils.sheet_to_json` function to convert your SpreadJS's workbook to JSON format and stored it as JSON file for this script to read.

### Example

```tsx
const { SpreadJSON } = require('./SpreadJSON');
let workbook = SpreadJSON.openByName("users", "workbook");
// In this situation, `users` is database folder and `workbook` is JSON file name.
// This path, you can customize it to suit your situation directly in class `SpreadJSON`
let sheet = workbook.getSheetByName("sheet1");

// What you can do:
workbook.getLastRow()

sheet.getRange("B3:D100");
sheet.getRange("A10").offset(0, 0, 100, 5);
sheet.getRange(1, 50);
sheet.getRange(1, 50, 100);
let range = sheet.getRange(1, 50, 100, 2);

range.getValues()
range.getCell(1, 5).getValue()

range.getCell(2, 1).setValue("Bob")
range.getCell(2, 1).setValues(["Bob"])
sheet.getRange(1, 50, 2, 2).setValues([
  ["Alice", 5],
  ["Demond", 10]
]);
```
