const XLSX = require('xlsx');
const path = require('path');

module.exports = async (req, res) => {
    try {
        const filename = req.query.file || 'Moodle_datein.xlsx';
        
        // Read the Excel file
        const excelPath = path.join(process.cwd(), filename);
        const workbook = XLSX.readFile(excelPath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Get the range of the sheet
        const range = XLSX.utils.decode_range(sheet['!ref']);
        
        // Get headers from the first row
        const headers = [];
        for(let C = range.s.c; C <= range.e.c; ++C) {
            const cell = sheet[XLSX.utils.encode_cell({r: 0, c: C})];
            headers[C] = cell ? cell.v : undefined;
        }
        
        // Convert the rest of the data to JSON using the headers
        const jsonData = [];
        for(let R = range.s.r + 1; R <= range.e.r; ++R) {
            const row = {};
            for(let C = range.s.c; C <= range.e.c; ++C) {
                const cell = sheet[XLSX.utils.encode_cell({r: R, c: C})];
                if (cell && headers[C]) {
                    row[headers[C]] = cell.v;
                }
            }
            if (Object.keys(row).length > 0) {
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