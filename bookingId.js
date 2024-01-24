import puppeteer from "puppeteer-extra";
import StealthPlugin from "puppeteer-extra-plugin-stealth";
import ExcelJS from "exceljs";

puppeteer.use(StealthPlugin());

(async () => {
  try {
    const browser = await puppeteer.launch({
      headless: false,
    });

    // Read the URL list from an Excel file
    const excelFilePath = "hotel1.xlsx"; // Replace with your actual file path
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);
    const worksheet = workbook.getWorksheet(1); // Replace with the actual sheet name
    const columnA = worksheet.getColumn("A").values.slice(2);

    const data = [];

    for (const url of columnA) {
      const page = await browser.newPage();
      await page.goto(url);
      await page.waitForSelector("div#wrap-hotelpage-top");

      // Extract the data from the current page
      const pageData = await page.evaluate(() => {
        const link = window.location.href;
        const title = document.title;
        const addressElement = document.querySelector(
          "span.hp_address_subtitle.js-hp_address_subtitle.jq_tooltip"
        );
        const address = addressElement ? addressElement.innerText : "";
      
        // Separate address into components
        const addressComponents = address.split(",");
      
        // Further split the postal code and city
        const postalCodeAndCity = addressComponents[1] ? addressComponents[1].trim().split(/\s+/) : [];
        
        // Extract postal code and city
        const postalCode = postalCodeAndCity.length > 0 ? postalCodeAndCity[0] : "";
        const city = postalCodeAndCity.length > 1 ? postalCodeAndCity.slice(1).join(' ') : "";
        const state = addressComponents[2] ? addressComponents[2].trim() : "";
      
        return {
          link,
          title,
          address,
          postalCode,
          city,
          state,
        };
      });
      
      // });
      data.push(pageData);
      await page.close();
    }

    await browser.close();

    // Save the data to an Excel file
    const outputWorkbook = new ExcelJS.Workbook();
    const outputWorksheet = outputWorkbook.addWorksheet("Sheet 1");

    // Add headers
    outputWorksheet.columns = [
      { header: "Link", key: "link" },
      { header: "Title", key: "title" },
      { header: "Address", key: "address" },
      { header: "Postal Code", key: "postalCode" },
      { header: "City", key: "city" },
      { header: "State", key: "state" },
    ];

    // Add data
    data.forEach((entry) => {
      outputWorksheet.addRow(entry);
    });

    // Save the output workbookc
    await outputWorkbook.xlsx.writeFile("list2-address.xlsx");
    console.timeEnd("Script Execution Time");
  } catch (error) {
    console.error(error);
  }
})();
