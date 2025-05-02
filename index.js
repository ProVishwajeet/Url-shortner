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

function createPersonalizedMessage(name, shortUrl) {
    // Get first name by splitting on space and taking first part
    const firstName = name.split(' ')[0];
    return `Hi ${firstName}, get ready for HCL SW Dubai SKO! Check out what MAX AI powered by Unica+ has coming up for you: ${shortUrl}`;
}

async function processExcelFile() {
    try {
        const inputFile = 'input.xlsx';
        console.log('Reading Excel file...');
        
        // Read the Excel file
        const workbook = XLSX.readFile(inputFile);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert to JSON
        const data = XLSX.utils.sheet_to_json(worksheet);
        
        let updatedCount = 0;
        
        // Process each row
        for (let row of data) {
            if (row.URL && !row['Short URL']) {
                console.log(`Processing URL for ${row.Name}...`);
                const shortUrl = await shortenUrl(row.URL);
                if (shortUrl) {
                    row['Short URL'] = shortUrl;
                    // Create and add personalized message
                    row['Message'] = createPersonalizedMessage(row.Name, shortUrl);
                    updatedCount++;
                    
                    // Update the file after each successful URL shortening
                    const newWorksheet = XLSX.utils.json_to_sheet(data);
                    workbook.Sheets[sheetName] = newWorksheet;
                    XLSX.writeFile(workbook, inputFile);
                    console.log(`✓ Updated short URL and message for ${row.Name}`);
                }
            } else if (row.URL && row['Short URL'] && !row['Message']) {
                // If URL and Short URL exist but no message, just add the message
                row['Message'] = createPersonalizedMessage(row.Name, row['Short URL']);
                updatedCount++;
                
                // Update the file
                const newWorksheet = XLSX.utils.json_to_sheet(data);
                workbook.Sheets[sheetName] = newWorksheet;
                XLSX.writeFile(workbook, inputFile);
                console.log(`✓ Added message for ${row.Name}`);
            }
        }
        
        console.log(`\nProcessing completed!`);
        console.log(`- Total records updated: ${updatedCount}`);
        console.log(`- File updated: ${inputFile}`);
        
    } catch (error) {
        console.error('Error processing Excel file:', error.message);
    }
}

// Run the script
processExcelFile(); 