const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const app = express();

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
        // Read the Excel file
        const excelPath = path.join(process.cwd(), 'Moodle_datein.xlsx');
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

// API endpoint to download Excel file
app.get('/api/download', async (req, res) => {
    try {
        const excelPath = path.join(process.cwd(), 'Moodle_datein.xlsx');
        
        // Check if file exists
        if (!fs.existsSync(excelPath)) {
            return res.status(404).json({ error: 'Excel file not found' });
        }

        // Set headers for file download
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Moodle_datein.xlsx');
        
        // Stream the file
        const fileStream = fs.createReadStream(excelPath);
        fileStream.pipe(res);
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

        // Save the file
        const excelPath = path.join(process.cwd(), 'Moodle_datein.xlsx');
        XLSX.writeFile(workbook, excelPath);

        res.json({ 
            message: 'File uploaded successfully',
            rowCount: jsonData.length
        });
    } catch (error) {
        console.error('Error uploading file:', error);
        res.status(500).json({ error: 'Failed to upload file' });
    }
});

// Handle root path
app.get('/', (req, res) => {
    res.json({ 
        message: 'Excel API is running',
        endpoints: {
            getData: '/api/excel-data',
            upload: '/api/upload',
            download: '/api/download'
        }
    });
});

// Export the Express API
module.exports = app; 