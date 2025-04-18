const XLSX = require('xlsx');

// Read the existing file
const workbook = XLSX.readFile('test_modified.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Create new data with modifications
const data = [
    { Name: 'Seminar_Begin', Data: 45710 },
    { Name: 'Seminar_Ende', Data: 45770 },
    { Name: 'Praktikum_Begin', Data: 45771 },
    { Name: 'Praktikum_ende', Data: 45833 },
    { Name: 'Test_Entry', Data: 45900 }
];

// Convert data to sheet
const newSheet = XLSX.utils.json_to_sheet(data);
workbook.Sheets[workbook.SheetNames[0]] = newSheet;

// Write the modified file
XLSX.writeFile(workbook, 'test_modified.xlsx');
console.log('File modified successfully'); 