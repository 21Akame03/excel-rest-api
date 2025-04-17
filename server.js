const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const app = express();
const port = process.env.PORT || 3000;

// Enable CORS for all routes
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    next();
});

// API endpoint to get Excel data as JSON
app.get('/api/excel-data', (req, res) => {
    try {
        // Read the Excel file
        const excelPath = path.join(__dirname, 'Moodle_datein.xlsx');
        const workbook = XLSX.readFile(excelPath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);
        
        // Send the JSON data
        res.json(jsonData);
    } catch (error) {
        console.error('Error processing Excel file:', error);
        res.status(500).json({ error: 'Failed to process Excel file' });
    }
});

// Handle root path
app.get('/', (req, res) => {
    res.json({ message: 'Excel API is running. Use /api/excel-data to get the Excel data.' });
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
    console.log(`API endpoint: http://localhost:${port}/api/excel-data`); 
}); 