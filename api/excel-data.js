const XLSX = require('xlsx');
const https = require('https');
const path = require('path');

const FILE_URLS = {
    'PublicMoodleNewsfeed.xlsx': 'https://digilern.hs-duesseldorf.de/cloud/s/bBZBbH6r8aTLMoy/download/PublicMoodleNewsfeed.xlsx',
    'PublicMoodleData.xlsx': 'https://digilern.hs-duesseldorf.de/cloud/s/6LC7Q982HJt28qi/download/PublicMoodleData.xlsx'
};

async function fetchExcelFile(url) {
    return new Promise((resolve, reject) => {
        https.get(url, response => {
            const chunks = [];
            response.on('data', chunk => chunks.push(chunk));
            response.on('end', () => resolve(Buffer.concat(chunks)));
            response.on('error', reject);
        }).on('error', reject);
    });
}

module.exports = async (req, res) => {
    try {
        const filename = req.query.file || 'Moodle_datein.xlsx';
        let workbook;

        // Check if we need to fetch from URL
        if (FILE_URLS[filename]) {
            const buffer = await fetchExcelFile(FILE_URLS[filename]);
            workbook = XLSX.read(buffer, { type: 'buffer' });
        } else {
            // Read local file as fallback
            const excelPath = path.join(process.cwd(), filename);
            workbook = XLSX.readFile(excelPath);
        }

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(sheet['!ref']);
        
        // Get headers from the first row (these will be your table headers)
        const headers = [];
        for(let C = range.s.c; C <= range.e.c; ++C) {
            const cell = sheet[XLSX.utils.encode_cell({r: 0, c: C})];
            headers[C] = cell ? cell.v : undefined;
        }
        
        // Convert the rest of the data to JSON using the headers
        const jsonData = [];
        for(let R = range.s.r + 1; R <= range.e.r; ++R) {
            const row = {};
            let hasData = false;
            for(let C = range.s.c; C <= range.e.c; ++C) {
                const cell = sheet[XLSX.utils.encode_cell({r: R, c: C})];
                if (cell && headers[C]) {
                    row[headers[C]] = cell.v;
                    hasData = true;
                }
            }
            if (hasData) {
                jsonData.push(row);
            }
        }
        
        // Set CORS headers
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
        
        // Send the JSON data
        res.json({
            filename: filename,
            headers: headers.filter(h => h !== undefined),
            data: jsonData
        });
    } catch (error) {
        console.error('Error processing Excel file:', error);
        res.status(500).json({ error: `Failed to process Excel file: ${error.message}` });
    }
};