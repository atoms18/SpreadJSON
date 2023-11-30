
# SpreadJSON

Did you ever working on your `Google Sheets/Apps Script` And using those `getRange`, `setValue`, `setValues` and etc?


This JS script will help you to do that directly on `Node.js`.


By `SpreadJSON` are designed to use a popular `SheetJS` Node.js library to stored your sheet in the format of JSON and provide a methods that you familiar.

So this script simply make you to be able to stored your sheet in the JSON format. And read or write to the sheet directly on Node.js.

(It helps you a lot if you have Apps Script project and want to migrate to Node.js to run faster and wherever you want!)

Caution: This project still not complete. There's a lot of features to add to make it more general and easy to use.

> This script use 3 dependencies which is `SheetJS` `fs` and `rwlock`.<br>
> Rwlock was used to implement an easy version of solution due to race condition problem.

## Prerequisites

This project requires NodeJS.
[Node](http://nodejs.org/) are really easy to install.
To make sure you have them available on your machine,
try running the following command.

```sh
$ node -v
v20.10.0
```

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

Caution:
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
