# URL Shortener with Excel Integration

This Node.js application reads URLs from an Excel file, shortens them using the Bitly API, and saves the results back to Excel.

## Features
- Read URLs from Excel file
- Shorten URLs using Bitly API
- Support for custom domains (with Bitly paid plans)
- Progress tracking and auto-save
- Batch processing support

## Setup
1. Clone the repository
2. Install dependencies:
```bash
npm install
```
3. Create a `.env` file with your Bitly token:
```
BITLY_ACCESS_TOKEN=your_bitly_token_here
```

## Input Excel Format
Create `input.xlsx` with these columns:
- Name
- Mobile Number
- URL
- Short URL

## Usage
```bash
node index.js
```

The script will:
1. Read from `input.xlsx`
2. Process each URL
3. Create `output.xlsx` with shortened URLs

## Rate Limits
- Free Plan: 1,000 links/month, 50 API calls/minute
- Paid Plans: Higher limits available

## Dependencies
- xlsx: Excel file handling
- axios: HTTP requests
- dotenv: Environment variables 