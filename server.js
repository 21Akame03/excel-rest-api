const express = require('express');
const XLSX = require('xlsx');
const multer = require('multer');
const app = express();

// In-memory storage for the Excel file
let excelBuffer = null;

// Initialize with the local Excel file
try {
    const workbook = XLSX.readFile('Moodle_datein.xlsx');
    excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    console.log('Initialized with local Excel file');
} catch (error) {
    console.error('Error reading local Excel file:', error);
}

// Configure multer for handling file uploads
const storage = multer.memoryStorage();
const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        console.log('Received file:', file.originalname, 'mimetype:', file.mimetype);
        if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
            file.mimetype === 'application/vnd.ms-excel' ||
            file.mimetype === 'application/octet-stream') {  // Added for binary uploads
            cb(null, true);
        } else {
            cb(new Error('Only Excel files are allowed'));
        }
    }
}).single('file');

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
            return res.status(404).json({ error: 'No Excel file available. Please upload a file first.' });
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
            return res.status(404).json({ error: 'No Excel file available. Please upload a file first.' });
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
app.post('/api/upload', (req, res) => {
    upload(req, res, async (err) => {
        try {
            if (err) {
                console.error('Multer error:', err);
                return res.status(400).json({ error: err.message });
            }

            if (!req.file) {
                console.error('No file received');
                return res.status(400).json({ error: 'No file uploaded' });
            }

            console.log('File received:', req.file.originalname, 'size:', req.file.size);

            // Read the uploaded file
            const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
            
            // Validate the Excel file
            if (workbook.SheetNames.length === 0) {
                console.error('No sheets in workbook');
                return res.status(400).json({ error: 'Excel file has no sheets' });
            }

            // Convert to JSON to validate data
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet);

            if (jsonData.length === 0) {
                console.error('No data in sheet');
                return res.status(400).json({ error: 'Excel file has no data' });
            }

            console.log('Parsed data:', JSON.stringify(jsonData));

            // Store the file buffer in memory
            excelBuffer = req.file.buffer;

            res.json({ 
                message: 'File uploaded successfully',
                rowCount: jsonData.length,
                data: jsonData
            });
        } catch (error) {
            console.error('Error processing upload:', error);
            res.status(500).json({ error: 'Failed to process uploaded file: ' + error.message });
        }
    });
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
                .info {
                    color: #0dcaf0;
                    text-align: center;
                    padding: 20px;
                    font-style: italic;
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
                <div id="info" class="info" style="display: none;">No data available. Please upload an Excel file.</div>
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
                            if (response.status === 404) {
                                document.getElementById('loading').style.display = 'none';
                                document.getElementById('info').style.display = 'block';
                                return;
                            }
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
                    document.getElementById('error').style.display = 'none';
                    document.getElementById('info').style.display = 'none';

                    if (data.length === 0) {
                        document.getElementById('info').style.display = 'block';
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

                    document.getElementById('loading').style.display = 'block';
                    document.getElementById('loading').textContent = 'Uploading file...';
                    document.getElementById('error').style.display = 'none';
                    document.getElementById('info').style.display = 'none';

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

                        // Display success message
                        document.getElementById('loading').textContent = 'File uploaded successfully!';
                        setTimeout(() => {
                            document.getElementById('loading').style.display = 'none';
                        }, 2000);

                        // Refresh the data display
                        displayData(result.data);
                    } catch (error) {
                        document.getElementById('loading').style.display = 'none';
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