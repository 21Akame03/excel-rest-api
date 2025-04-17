const fs = require('fs');
const path = require('path');

module.exports = (req, res) => {
    if (req.url === '/') {
        // Serve the HTML file
        const htmlPath = path.join(process.cwd(), 'public', 'index.html');
        const htmlContent = fs.readFileSync(htmlPath, 'utf8');
        
        res.setHeader('Content-Type', 'text/html');
        res.send(htmlContent);
    } else {
        // For other routes, return API info
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
        res.json({ message: 'Excel API is running. Use /api/excel-data to get the Excel data.' });
    }
}; 