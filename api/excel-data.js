import XLSX from 'xlsx';
import https from 'https';
import path from 'path';

const FILE_URLS = {
    'PublicMoodleNewsfeed.xlsx': 'https://digilern.hs-duesseldorf.de/cloud/s/bBZBbH6r8aTLMoy/download/PublicMoodleNewsfeed.xlsx',
    'PublicMoodleData.xlsx': 'https://digilern.hs-duesseldorf.de/cloud/s/6LC7Q982HJt28qi/download/PublicMoodleData.xlsx'
};

async function fetchExcelFile(url) {
    return new Promise((resolve, reject) => {
        const agent = new https.Agent({
            rejectUnauthorized: false
        });
        
        https.get(url, { agent }, response => {
            if (response.statusCode === 302 || response.statusCode === 301) {
                // Handle redirect
                https.get(response.headers.location, { agent }, redirectedResponse => {
                    const chunks = [];
                    redirectedResponse.on('data', chunk => chunks.push(chunk));
                    redirectedResponse.on('end', () => resolve(Buffer.concat(chunks)));
                    redirectedResponse.on('error', reject);
                }).on('error', reject);
            } else {
                const chunks = [];
                response.on('data', chunk => chunks.push(chunk));
                response.on('end', () => resolve(Buffer.concat(chunks)));
                response.on('error', reject);
            }
        }).on('error', reject);
    });
}

export default async function handler(req, res) {
    // Enable CORS
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,POST,PUT');
    res.setHeader(
        'Access-Control-Allow-Headers',
        'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
    );

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        const filename = req.query.file;
        if (!filename || !FILE_URLS[filename]) {
            return res.status(400).json({ error: 'Please specify a valid file: PublicMoodleNewsfeed.xlsx or PublicMoodleData.xlsx' });
        }

        console.log('Fetching from URL:', FILE_URLS[filename]);
        const buffer = await fetchExcelFile(FILE_URLS[filename]);
        const workbook = XLSX.read(buffer, { type: 'buffer' });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Convert to JSON with specific options
        const jsonData = XLSX.utils.sheet_to_json(sheet, {
            raw: true,
            defval: null,
            header: 1
        });

        // Extract headers from first row
        const headers = jsonData[0];
        
        // Process remaining rows based on file type
        let data;
        if (filename === 'PublicMoodleData.xlsx') {
            // For PublicMoodleData.xlsx, use specific headers
            const expectedHeaders = ['Fach', 'Versuch', 'Variable', 'Daten'];
            const processedData = jsonData.slice(1)
                .filter(row => row.some(cell => cell != null)) // Remove empty rows
                .map(row => {
                    const rowData = {};
                    expectedHeaders.forEach((header, index) => {
                        rowData[header] = row[index] !== undefined ? row[index] : null;
                    });
                    return rowData;
                })
                // Filter out duplicate rows based on all fields combined
                .filter((row, index, self) => 
                    index === self.findIndex(r => 
                        r.Fach === row.Fach && 
                        r.Versuch === row.Versuch && 
                        r.Variable === row.Variable && 
                        r.Daten === row.Daten
                    )
                );
            return {
                filename,
                headers: expectedHeaders,
                data: processedData
            };
        } else {
            // For other files (like PublicMoodleNewsfeed.xlsx), use existing processing
            data = jsonData.slice(1)
                .filter(row => row.some(cell => cell != null)) // Remove empty rows
                .map(row => {
                    const rowData = {};
                    headers.forEach((header, index) => {
                        if (header) {
                            rowData[header] = row[index] !== undefined ? row[index] : null;
                        }
                    });
                    return rowData;
                });
        }

        res.status(200).json({
            filename: filename,
            headers: headers.filter(Boolean),
            data: data
        });
    } catch (error) {
        console.error('Error processing Excel file:', error);
        res.status(500).json({ error: `Failed to process Excel file: ${error.message}` });
    }
}