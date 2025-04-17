const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const app = express();
const port = 3000;

// Serve static files from the current directory
app.use(express.static(__dirname));

// API endpoint to get Excel data as JSON
app.get('/api/excel-data', (req, res) => {
    try {
        // Read the Excel file
        const workbook = XLSX.readFile('Moodle_datein.xlsx');
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);
        
        // Send the JSON data
        res.json(jsonData);
    } catch (error) {
        console.error('Error processing Excel file:', error);
        res.status(500).json({ error: 'Failed to process Excel file' });
    }
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
}); 