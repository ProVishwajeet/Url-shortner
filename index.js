require('dotenv').config();
const XLSX = require('xlsx');
const axios = require('axios');

// Bitly API configuration
const BITLY_API_URL = 'https://api-ssl.bitly.com/v4/shorten';
const BITLY_ACCESS_TOKEN = process.env.BITLY_ACCESS_TOKEN;

async function shortenUrl(longUrl) {
    try {
        const response = await axios.post(
            BITLY_API_URL,
            { long_url: longUrl },
            {
                headers: {
                    'Authorization': `Bearer ${BITLY_ACCESS_TOKEN}`,
                    'Content-Type': 'application/json'
                }
            }
        );
        return response.data.link;
    } catch (error) {
        console.error(`Error shortening URL ${longUrl}:`, error.message);
        return null;
    }
}

async function processExcelFile() {
    try {
        // Read the Excel file
        const workbook = XLSX.readFile('input.xlsx');
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert to JSON
        const data = XLSX.utils.sheet_to_json(worksheet);
        
        // Process each row
        for (let row of data) {
            if (row.URL && !row['Short URL']) {
                console.log(`Processing URL for ${row.Name}...`);
                const shortUrl = await shortenUrl(row.URL);
                if (shortUrl) {
                    row['Short URL'] = shortUrl;
                }
            }
        }
        
        // Convert back to worksheet
        const newWorksheet = XLSX.utils.json_to_sheet(data);
        
        // Create new workbook and save
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
        XLSX.writeFile(newWorkbook, 'output.xlsx');
        
        console.log('Processing completed! Check output.xlsx for results.');
    } catch (error) {
        console.error('Error processing Excel file:', error.message);
    }
}

// Run the script
processExcelFile(); 