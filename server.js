const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const fetch = require('node-fetch');
const app = express();
const port = process.env.PORT || 3000;

// Enable CORS
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

// Serve static files from the current directory
app.use(express.static(__dirname));

// API endpoint to get Excel data as JSON
app.get('/api/excel-data', async (req, res) => {
    try {
        // GitHub raw content URL
        const excelUrl = "https://raw.githubusercontent.com/21Akame03/excel-rest-api/main/Moodle_datein.xlsx";
        
        // Fetch the Excel file
        const response = await fetch(excelUrl);
        if (!response.ok) {
            throw new Error(`Failed to fetch Excel file: ${response.statusText}`);
        }
        
        // Get the file as array buffer
        const arrayBuffer = await response.arrayBuffer();
        
        // Read the Excel file
        const workbook = XLSX.read(arrayBuffer, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);
        
        // Send the JSON data
        res.json(jsonData);
    } catch (error) {
        console.error('Error processing Excel file:', error);
        res.status(500).json({ 
            error: 'Failed to process Excel file',
            details: error.message 
        });
    }
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ status: 'ok' });
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
}); 