const xlsx = require('xlsx');
const path = require('path');
const assets = require('./assets/__assets.js');
// create a xlsx file
const workbook = xlsx.utils.book_new();

let student_data = [
  {
    Student: 'Nikhil',
    Age: 22,
    Branch: 'ISE',
    Marks: 70,
  },
  {
    Student: 'Amitha',
    Age: 21,
    Branch: 'EC',
    Marks: 80,
  },
];

const ws = xlsx.utils.json_to_sheet(student_data);

xlsx.utils.book_append_sheet(workbook, ws, 'Sheet1');

xlsx.writeFile(workbook, path.resolve(assets, 'data.xlsx'));
