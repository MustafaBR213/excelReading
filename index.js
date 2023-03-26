// const XLSX = require('xlsx');
// const workbook = XLSX.readFile('eee.xlsx');
// const sheetName = workbook.SheetNames[0];
// const worksheet = workbook.Sheets[sheetName];

// // Define cell ranges for the columns A, B, and C
// const rangeA = XLSX.utils.decode_range(worksheet['!ref']);
// rangeA.s.col = 0; // Start column
// rangeA.e.col = 0; // End column
// const rangeB = XLSX.utils.decode_range(worksheet['!ref']);
// rangeB.s.col = 1; // Start column
// rangeB.e.col = 1; // End column
// const rangeC = XLSX.utils.decode_range(worksheet['!ref']);
// rangeC.s.col = 2; // Start column
// rangeC.e.col = 2; // End column

// // Loop through rows and sum values in columns A and B, and save result in column C
// for (let i = rangeA.s.r; i <= rangeA.e.r; i++) {
//   const cellA = worksheet[XLSX.utils.encode_cell({r: i, c: rangeA.s.col})];
//   const cellB = worksheet[XLSX.utils.encode_cell({r: i, c: rangeB.s.col})];
//   const sum = (cellA ? parseInt(cellA.v) : 0) * (cellB ? parseInt(cellB.v) : 0);
//   const cellC = {t: 'n', v: sum};
//   worksheet[XLSX.utils.encode_cell({r: i, c: rangeC.s.col})] = cellC;
// }

// // Save the modified workbook to a new file
// XLSX.writeFile(workbook, 'newt.xlsx');




// const XLSX = require('xlsx');
// const workbook = XLSX.readFile('fer.xlsx');
// const sheetName = workbook.SheetNames[0];
// const worksheet = workbook.Sheets[sheetName];

// // Define column names for the columns A, B, and C
// const columnNameA = 'h=Hdr-m-d, m';
// const columnNameB = 'K, m/d';
// const columnNameC = 'asd';

// // Get the range of cells for the columns A, B, and C by column name
// const rangeA = XLSX.utils.decode_range(worksheet['!ref']);
// rangeA.s.col = XLSX.utils.decode_col(columnNameA);
// rangeA.e.col = rangeA.s.col;
// const rangeB = XLSX.utils.decode_range(worksheet['!ref']);
// rangeB.s.col = XLSX.utils.decode_col(columnNameB);
// rangeB.e.col = rangeB.s.col;
// const rangeC = XLSX.utils.decode_range(worksheet['!ref']);
// rangeC.s.col = XLSX.utils.decode_col(columnNameC);
// rangeC.e.col = rangeC.s.col;

// // Loop through rows and sum values in columns A and B, and save result in column C
// for (let i = rangeA.s.r; i <= rangeA.e.r; i++) {
//   const cellA = worksheet[XLSX.utils.encode_cell({r: i, c: rangeA.s.col})];
//   const cellB = worksheet[XLSX.utils.encode_cell({r: i, c: rangeB.s.col})];
//   const sum = (cellA ? parseInt(cellA.v) : 1) * (cellB ? parseInt(cellB.v) : 1);
//   const cellC = {t: 'n', v: sum};
//   worksheet[XLSX.utils.encode_cell({r: i, c: rangeC.s.col})] = cellC;
// }

// // Save the modified workbook to a new file
// XLSX.writeFile(workbook, 'newssw.xlsx');



const XLSX = require('xlsx');
const workbook = XLSX.readFile('fer.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Define column names for the columns A, B, and C
const columnNameA = 'h=Hdr-m-d, m';
const columnNameB = 'K, m/d';
const columnNameC = 'asd';

// Get the range of cells for the columns A, B, and C by column name
const rangeA = XLSX.utils.decode_range(worksheet['!ref']);
rangeA.s.col = XLSX.utils.decode_col(columnNameA);
rangeA.e.col = rangeA.s.col;
const rangeB = XLSX.utils.decode_range(worksheet['!ref']);
rangeB.s.col = XLSX.utils.decode_col(columnNameB);
rangeB.e.col = rangeB.s.col;
const rangeC = XLSX.utils.decode_range(worksheet['!ref']);
rangeC.s.col = XLSX.utils.decode_col(columnNameC);
rangeC.e.col = rangeC.s.col;

// Loop through rows and sum values in columns A and B, and save result in column C
for (let i = rangeA.s.r; i <= rangeA.e.r; i++) {
  const cellA = worksheet[XLSX.utils.encode_cell({r: i, c: rangeA.s.col})];
  const cellB = worksheet[XLSX.utils.encode_cell({r: i, c: rangeB.s.col})];
  const sum = (cellA ? parseInt(cellA.v) : 0) + (cellB ? parseInt(cellB.v) : 0);
  const cellC = {t: 'n', v: sum};
  worksheet[XLSX.utils.encode_cell({r: i, c: rangeC.s.col})] = cellC;
}

// Save the modified workbook to a new file
XLSX.writeFile(workbook, 'news.xlsx');
