const express = require('express');
const XLSX = require('xlsx');
const multer = require('multer');
const app = express();

// In-memory storage for the Excel file
let excelBuffer = null;

// Initialize with default data if no file is uploaded
try {
    const defaultData = [
        { Name: "Seminar_Begin", Data: 45709 },
        { Name: "Seminar_Ende", Data: 45768 },
        { Name: "Praktikum_Begin", Data: 45769 },
        { Name: "Praktikum_ende", Data: 45831 }
    ];
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(defaultData);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
} catch (error) {
    console.error('Error initializing default data:', error);
}

// Configure multer for handling file uploads
const storage = multer.memoryStorage();
const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
            file.mimetype === 'application/vnd.ms-excel') {
            cb(null, true);
        } else {
            cb(new Error('Only Excel files are allowed'));
        }
    }
});

// Enable CORS for all routes
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    if (req.method === 'OPTIONS') {
        return res.sendStatus(200);
    }
    next();
});

// Parse JSON bodies
app.use(express.json());

// API endpoint to get Excel data as JSON
app.get('/api/excel-data', async (req, res) => {
    try {
        if (!excelBuffer) {
            return res.status(404).json({ error: 'No Excel file available' });
        }

        const workbook = XLSX.read(excelBuffer, { type: 'buffer' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);
        
        res.json(jsonData);
    } catch (error) {
        console.error('Error processing Excel file:', error);
        res.status(500).json({ error: 'Failed to process Excel file' });
    }
});

// API endpoint to download Excel file
app.get('/api/download', async (req, res) => {
    try {
        if (!excelBuffer) {
            return res.status(404).json({ error: 'No Excel file available' });
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Moodle_datein.xlsx');
        res.send(excelBuffer);
    } catch (error) {
        console.error('Error downloading file:', error);
        res.status(500).json({ error: 'Failed to download file' });
    }
});

// API endpoint to upload Excel file
app.post('/api/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        // Read the uploaded file
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        
        // Validate the Excel file
        if (workbook.SheetNames.length === 0) {
            return res.status(400).json({ error: 'Excel file has no sheets' });
        }

        // Convert to JSON to validate data
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        if (jsonData.length === 0) {
            return res.status(400).json({ error: 'Excel file has no data' });
        }

        // Store the file buffer in memory
        excelBuffer = req.file.buffer;

        res.json({ 
            message: 'File uploaded successfully',
            rowCount: jsonData.length
        });
    } catch (error) {
        console.error('Error uploading file:', error);
        res.status(500).json({ error: 'Failed to upload file' });
    }
});

// Serve static HTML for the root path
app.get('/', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Excel Data Viewer</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    margin: 20px;
                    background-color: #f5f5f5;
                }
                .container {
                    max-width: 1200px;
                    margin: 0 auto;
                    background-color: white;
                    padding: 20px;
                    border-radius: 8px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                }
                h1 {
                    color: #333;
                    text-align: center;
                    margin-bottom: 30px;
                }
                .buttons {
                    display: flex;
                    justify-content: center;
                    gap: 20px;
                    margin-bottom: 30px;
                }
                .button {
                    padding: 10px 20px;
                    background-color: #007bff;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                    text-decoration: none;
                }
                .button:hover {
                    background-color: #0056b3;
                }
                #fileInput {
                    display: none;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 20px;
                }
                th, td {
                    padding: 12px;
                    text-align: left;
                    border-bottom: 1px solid #ddd;
                }
                th {
                    background-color: #f8f9fa;
                    font-weight: bold;
                }
                tr:hover {
                    background-color: #f5f5f5;
                }
                .loading {
                    text-align: center;
                    padding: 20px;
                    font-style: italic;
                    color: #666;
                }
                .error {
                    color: #dc3545;
                    text-align: center;
                    padding: 20px;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Excel Data Viewer</h1>
                <div class="buttons">
                    <input type="file" id="fileInput" accept=".xlsx,.xls">
                    <button class="button" onclick="document.getElementById('fileInput').click()">Upload New Excel File</button>
                    <a href="/api/download" class="button">Download Current Excel</a>
                </div>
                <div id="loading" class="loading">Loading data...</div>
                <div id="error" class="error" style="display: none;"></div>
                <table id="dataTable">
                    <thead>
                        <tr id="tableHeader"></tr>
                    </thead>
                    <tbody id="tableBody"></tbody>
                </table>
            </div>

            <script>
                async function fetchData() {
                    try {
                        const response = await fetch('/api/excel-data');
                        if (!response.ok) {
                            throw new Error('Failed to fetch data');
                        }
                        const data = await response.json();
                        displayData(data);
                    } catch (error) {
                        document.getElementById('loading').style.display = 'none';
                        document.getElementById('error').style.display = 'block';
                        document.getElementById('error').textContent = 'Error loading data: ' + error.message;
                    }
                }

                function displayData(data) {
                    const tableHeader = document.getElementById('tableHeader');
                    const tableBody = document.getElementById('tableBody');
                    
                    // Clear existing content
                    tableHeader.innerHTML = '';
                    tableBody.innerHTML = '';
                    document.getElementById('loading').style.display = 'none';

                    if (data.length === 0) {
                        document.getElementById('error').style.display = 'block';
                        document.getElementById('error').textContent = 'No data available';
                        return;
                    }

                    // Create table headers
                    const headers = Object.keys(data[0]);
                    headers.forEach(header => {
                        const th = document.createElement('th');
                        th.textContent = header;
                        tableHeader.appendChild(th);
                    });

                    // Create table rows
                    data.forEach(row => {
                        const tr = document.createElement('tr');
                        headers.forEach(header => {
                            const td = document.createElement('td');
                            td.textContent = row[header] || '';
                            tr.appendChild(td);
                        });
                        tableBody.appendChild(tr);
                    });
                }

                // Handle file upload
                document.getElementById('fileInput').addEventListener('change', async (event) => {
                    const file = event.target.files[0];
                    if (!file) return;

                    const formData = new FormData();
                    formData.append('file', file);

                    try {
                        const response = await fetch('/api/upload', {
                            method: 'POST',
                            body: formData
                        });

                        const result = await response.json();
                        if (!response.ok) {
                            throw new Error(result.error || 'Upload failed');
                        }

                        // Refresh the data display
                        fetchData();
                    } catch (error) {
                        document.getElementById('error').style.display = 'block';
                        document.getElementById('error').textContent = 'Upload error: ' + error.message;
                    }
                });

                // Fetch data when the page loads
                fetchData();
            </script>
        </body>
        </html>
    `);
});

// Export the Express API
module.exports = app; 