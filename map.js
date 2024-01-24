import puppeteer from "puppeteer-extra";
import StealthPlugin from "puppeteer-extra-plugin-stealth"
import ExcelJS from "exceljs"
import path from "path"
import * as XLSX from "xlsx"
import AdblockerPlugin from "puppeteer-extra-plugin-adblocker"

puppeteer.use(StealthPlugin());
puppeteer.use(AdblockerPlugin({ blockTrackers: true }));

(async ()=> {
    try {
        const browser = await puppeteer.launch({
            headless: 'new',
        })
    
        const url = "https://www.google.com/maps"
    
        const page = await browser.newPage()
        await page.goto(url)

        // Make new object excel
        
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('./hotel-list/hotel1-address.xlsx');
        const worksheet = workbook.getWorksheet(1);

        const columnA = worksheet.getColumn('B').values.slice(2);

        // Wait for search selector
        await page.waitForSelector('input#searchboxinput')
        // // const input = await page.$('input#searchboxinput')
        const data = []

        let hotelCount = 0;
        // const defaultDelay = 1000;
        for(const name of columnA){
            try {
                console.log(`Processing hotel ${hotelCount + 1}: ${name}`);
                hotelCount++; 
                await page.click('input#searchboxinput', {clickCount: 3})

                await page.type('input#searchboxinput', name, {delay: 50})
                await page.keyboard.press('Enter')

                await page.waitForTimeout(8000)
                // await page.waitForNavigation({waitUntil: 'domcontentloaded'})

                // Extract data from the page
                const result = await page.evaluate((name) => {
                    const phoneElement = document.querySelector('div:nth-child(5) > button > div > div.rogA2c > div.Io6YTe.fontBodyMedium.kR99db');
                    const websiteElement = document.querySelector('a.CsEnBe');
                    
                    const phone = phoneElement ? phoneElement.textContent?.trim() : '';
                    const website = websiteElement ? websiteElement.href : '';
                    return { name, phone, website };
                }, name);

                data.push(result)
                console.log(result)
                // console.log(result)
            } catch (error) {
                console.log(error)
            }
        }   
        const outputDirectory = "./"; // Relative path to the 'coffe' directory
        const currentModuleURL = import.meta.url;
        const currentModulePath = new URL(currentModuleURL).pathname;
        const outputFilePath = path.join(path.dirname(currentModulePath), outputDirectory, 'maps-hotel1.xlsx');

        // Save the data for this subCategory to an Excel file
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
        XLSX.writeFile(wb, outputFilePath);
    
        await browser.close()    
    } catch (error) {
        console.log(error)
    }
})()