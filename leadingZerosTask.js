var excel = require('excel4node');
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');
var style = workbook.createStyle({
  font: {
    color: '#000000',
    size: 12
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -'
});

const records = [
      {
        empID : "00123",
        empName : "Neha",
        salary : 20000
        },
      {
        empID : "00234",
        empName : "Sushma",
        salary : 10000
      },
      {
        empID : "11345",
        empName : "Harika",
        salary : 15000
      }
    ];

const headersCount = Object.keys(records[0]).length;
const headers = Object.keys(records[0]);
var i;
//to set headers
for(i = 0; i < headersCount; i++) {
    worksheet.cell(1, i + 1).string(String(headers[i])).style(style);
}
var j;
//to set values of headers
for(i = 0; i < headersCount; i++) {
    for(j = 0; j < records.length; j++) {
        worksheet.cell(j + 2, i + 1).string(String(records[j][headers[i]])).style(style);
    }
}

workbook.write('outputTaskA.xlsx');
