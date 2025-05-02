require('dotenv').config();
const XLSX = require('xlsx');
const axios = require('axios');

// Bitly API configuration
const BITLY_API_URL = 'https://api-ssl.bitly.com/v4/bitlinks';
const BITLY_ACCESS_TOKEN = process.env.BITLY_ACCESS_TOKEN;
const CUSTOM_DOMAIN = 'hclsftw.co';

async function shortenUrl(longUrl, firstName) {
    try {
        // Convert firstName to lowercase and remove any special characters
        const customBackHalf = firstName ? firstName.toLowerCase().replace(/[^a-z0-9]/g, '') : '';
        const payload = {
            long_url: longUrl,
            domain: CUSTOM_DOMAIN,
            title: `Link for ${firstName}`
        };
        if (customBackHalf) {
            payload.custom_bitlink = `${CUSTOM_DOMAIN}/${customBackHalf}`;
        }
        const response = await axios.post(
            BITLY_API_URL,
            payload,
            {
                headers: {
                    'Authorization': `Bearer ${BITLY_ACCESS_TOKEN}`,
                    'Content-Type': 'application/json'
                }
            }
        );
        return response.data.link;
    } catch (error) {
        if (error.response && error.response.status === 422) {
            // If custom name is taken, append a random number
            console.log(`Custom URL ${CUSTOM_DOMAIN}/${firstName} is already taken. Trying with a random suffix...`);
            const randomSuffix = Math.floor(Math.random() * 1000);
            try {
                const retryResponse = await axios.post(
                    BITLY_API_URL,
                    {
                        long_url: longUrl,
                        domain: CUSTOM_DOMAIN,
                        title: `Link for ${firstName}`,
                        custom_bitlink: `${CUSTOM_DOMAIN}/${customBackHalf}${randomSuffix}`
                    },
                    {
                        headers: {
                            'Authorization': `Bearer ${BITLY_ACCESS_TOKEN}`,
                            'Content-Type': 'application/json'
                        }
                    }
                );
                return retryResponse.data.link;
            } catch (retryError) {
                console.error(`Error creating custom URL with suffix: ${retryError.message}`);
                // Fallback to default with custom domain
                const defaultResponse = await axios.post(
                    BITLY_API_URL,
                    {
                        long_url: longUrl,
                        domain: CUSTOM_DOMAIN,
                        title: `Link for ${firstName}`
                    },
                    {
                        headers: {
                            'Authorization': `Bearer ${BITLY_ACCESS_TOKEN}`,
                            'Content-Type': 'application/json'
                        }
                    }
                );
                return defaultResponse.data.link;
            }
        }
        console.error(`Error shortening URL ${longUrl}:`, error.message);
        return null;
    }
}

function createPersonalizedMessage(firstName, shortUrl) {
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
            const firstName = row['First Name'] ? row['First Name'].trim() : '';
            const longUrl = row['Long URL'];
            if (longUrl && !row['Short URL']) {
                console.log(`Processing URL for ${firstName || row['Full Name'] || 'Unknown'}...`);
                const shortUrl = await shortenUrl(longUrl, firstName);
                if (shortUrl) {
                    row['Short URL'] = shortUrl;
                    // Create and add personalized message
                    row['Message'] = createPersonalizedMessage(firstName, shortUrl);
                    updatedCount++;
                    
                    // Update the file after each successful URL shortening
                    const newWorksheet = XLSX.utils.json_to_sheet(data);
                    workbook.Sheets[sheetName] = newWorksheet;
                    XLSX.writeFile(workbook, inputFile);
                    console.log(`✓ Updated short URL and message for ${firstName || row['Full Name'] || 'Unknown'}`);
                }
            } else if (longUrl && row['Short URL'] && !row['Message']) {
                // If Long URL and Short URL exist but no message, just add the message
                row['Message'] = createPersonalizedMessage(firstName, row['Short URL']);
                updatedCount++;
                
                // Update the file
                const newWorksheet = XLSX.utils.json_to_sheet(data);
                workbook.Sheets[sheetName] = newWorksheet;
                XLSX.writeFile(workbook, inputFile);
                console.log(`✓ Added message for ${firstName || row['Full Name'] || 'Unknown'}`);
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