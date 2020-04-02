const fs = require('fs');
const child_process = require('child_process');
const XLSX = require("xlsx");

class XLSXWorkbook {
    constructor(input) {
        if (input instanceof Buffer)
            this.loadWorkbookFromBuffer(input)
        else if (input instanceof String)
            this.loadWorkbookFromFile(input)
        else if (input instanceof URL)
            this.loadWorkbookFromURL(input)
        else if (typeof input == "undefined")
            return;
        else
            throw "XlsxWorkbook.constructor: bad input type";
    }

    loadWorkbookFromBuffer(data) {
        this.workbook = XLSX.read(data, {type: "buffer"});
    }
    loadWorkbookFromURL(url) {
        this.workbook = XLSX.read(child_process.execSync('curl -s ' + url), {type: "buffer"});
    }
    loadWorkbookFromFile(filename) {
        this.workbook = XLSX.readFile(filename);
    }
    getSpreadsheetAsObject(stylesheetName) {
        return XLSX.utils.sheet_to_json(this.workbook.Sheets[stylesheetName]);
    }
    getWorkbookAsObject() {
        const result = {};
        for (let sheetname of this.stylesheetNames()) {
            result[sheetname] = XLSX.utils.sheet_to_json(this.workbook.Sheets[sheetname]);
        }
        return result;
    }
    getWorkbook() {
        return this.workbook;
    }
    stylesheetNames() {
        return this.workbook.SheetNames;
    }
    saveWorkbookToFile(filename) {
        return fs.writeFileSync(filename, this.workbook);
    }
}

module.exports = XLSXWorkbook;
