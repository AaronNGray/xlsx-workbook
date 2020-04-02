const XLSXWorkbook = require("./XLSXWorkbook");

const dataURL = new URL("https://fingertips.phe.org.uk/documents/Historic%20COVID-19%20Dashboard%20Data.xlsx");
const filename = "Historic COVID-19 Dashboard Data.xlsx"

var workbook = new XLSXWorkbook(dataURL);

var data = workbook.getWorkbookAsObject();

console.log(JSON.stringify(data, 2, 2));

