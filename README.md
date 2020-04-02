# xlsx-workbook class

A wrapper based on https://github.com/SheetJS/sheetjs

## install

```
npm install --save xlsx-workbook
```

## Usage

```
const XLSXWorkbook = require("xlsx-workbook");

var wb = new XLSXWorkbook();

var spreadsheets = wb.getSpreadsheets
```
## API

```
class XLSXWorkbook

    constructor(Buffer|URL|String) throws "XlsxWorkbook.constructor: bad input type";

    loadWorkbookFromBuffer(Buffer)

    loadWorkbookFromURL(URL)

    loadWorkbookFromFile(String)

    getSpreadsheetNames():String[]

    getSpreadsheetAsObject(name:String)
    getWorkbookAsObject():Object

    getWorkbook()

    saveWorkbookToFile(filename:String|Buffer|URL|Integer)
```

## Example

### test.js
[test.js](https://github.com/AaronNGray/xlsx-workbook/blob/master/test.js) An example based on need.

```
const XLSXWorkbook = require("./index.js");

const dataURL = new URL("https://fingertips.phe.org.uk/documents/Historic%20COVID-19%20Dashboard%20Data.xlsx");
const filename = "Historic COVID-19 Dashboard Data.xlsx"

var workbook = new XLSXWorkbook(dataURL);

var data = workbook.getWorkbookAsObject();

console.log(JSON.stringify(data, 2, 2));

workbook.saveWorkbookToFile(filename);
```
