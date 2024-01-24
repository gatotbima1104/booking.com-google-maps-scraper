import fs from 'fs'
import puppeteer from 'puppeteer-extra';
import ExcelJS from "exceljs"
import StealthPlugin from "puppeteer-extra-plugin-stealth"

puppeteer.use(StealthPlugin());

async function extractEmails(url) {
    const browser = await puppeteer.launch({
        headless: "new"
    });
    const page = await browser.newPage();

    try {
        await page.goto(url, { waitUntil: 'domcontentloaded' });

        // Wait for a short duration to ensure dynamic content is loaded
        await page.waitForTimeout(2000);

        // Extract emails using a regular expression
        const emails = await page.evaluate(() => {
            const emailPattern = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
            const text = document.body.innerText;
            return text.match(emailPattern) || [];
        });

        return [...new Set(emails)]; // Remove duplicates using Set
    } catch (error) {
        console.error(`Failed to extract emails from ${url}. Error: ${error}`);
        return [];
    } finally {
        await browser.close();
    }
}

// Read the Excel file using exceljs
const excelFilePath = 'hotel.xlsx';
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(excelFilePath);

const worksheet = workbook.getWorksheet(1); // Assuming the URLs are in the first worksheet

// Create a new worksheet for storing results
const resultWorksheet = workbook.addWorksheet('EmailResults');
resultWorksheet.columns = [{ header: 'Website', key: 'website' }, { header: 'Emails', key: 'emails' }];

// Iterate over the URLs and extract emails
for (let index = 2; index <= worksheet.rowCount; index++) {
    const url = worksheet.getCell(`D${index}`).value;
    const emails = await extractEmails(url);
    
    console.log(emails);

    // Add the website and emails to the result worksheet
    resultWorksheet.addRow({ website: url, emails: emails.join(', ') });

    console.log(`Link number ${index - 1}: ${url}`);
}

// Save the result to a new Excel file
const resultExcelFile = 'result-with-js.xlsx';
await workbook.xlsx.writeFile(resultExcelFile);

console.log(`Extracted emails saved to ${resultExcelFile}`);
