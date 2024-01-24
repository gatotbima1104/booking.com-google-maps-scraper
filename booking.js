import puppeteer from "puppeteer-extra";
import StealthPlugin from "puppeteer-extra-plugin-stealth";
import ExcelJS from "exceljs";

puppeteer.use(StealthPlugin());

(async () => {
  try {
    const browser = await puppeteer.launch({
      headless: "new",
    });

    const baseUrl =
      "https://www.booking.com/searchresults.es.html?label=gen173nr-1BCAEoggI46AdIM1gEaEaIAQGYAQq4AQfIAQ3YAQHoAQGIAgGoAgO4Ap-y_6wGwAIB0gIkYzk4NmMzZDctN2FhZS00MTdiLWEzMDQtNGE4ZWJkNjI1OTVi2AIF4AIB&sid=8af25a1c3977a6764370f3db4f54f22b&aid=304142&ss=Sachsen-Anhalt%2C+Deutschland&ssne=Hamburgo&ssne_untouched=Hamburgo&lang=es&src=searchresults&dest_id=713&dest_type=region&ac_position=2&ac_click_type=b&ac_langcode=de&ac_suggestion_list_length=5&search_selected=true&search_pageview_id=a8c556e0581f0275&ac_meta=GhBhOGM1NTZlMDU4MWYwMjc1IAIoATICZGU6DlNhY2hzZW4tQW5oYWx0QABKAFAA&checkin=2024-04-11&checkout=2024-04-12&group_adults=1&no_rooms=1&group_children=0&nflt=privacy_type%3D3&offset=";

    const data = [];
    // let idCounter = 1; // Initialize ID counter

    for (let offset = 0; offset <= 525; offset += 25) {
      const page = await browser.newPage();
      await page.goto(`${baseUrl}${offset}`);
      await page.waitForTimeout(1000);

      await page.waitForSelector("div.d4924c9e74");

      //  extract the data from the current page
      const pageData = await page.evaluate(() => {
        const productElements = Array.from(
          document.querySelectorAll('div[aria-label="Alojamiento"]')
        );

        return productElements.map((product) => {
          const linkElement = product.querySelector(
            'a[data-testid="availability-cta-btn"]'
          );
          const titleElement = product.querySelector(
            'div[data-testid="title"]'
          );

          const link = linkElement ? linkElement.href : "";
          const title = titleElement
            ? titleElement.textContent?.trim().replace(/\n/g, " ")
            : "";
          return { link, title }; // Placeholder ID, actual ID will be set later
        });
      });

      // Set the correct IDs based on the counter
      // for (const entry of pageData) {
      //     entry.id = idCounter++;
      // }

      data.push(...pageData);
      // console.log(data)
      await page.close();
    }

    await browser.close();

    // Save the data to an Excel file
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    // Add headers
    worksheet.columns = [
      { header: "Link", key: "link" },
      { header: "Title", key: "title" },
    ];

    // Add data
    data.forEach((entry) => {
      worksheet.addRow(entry);
    });

    // Save the workbook
    await workbook.xlsx.writeFile("list37.xlsx");
  } catch (error) {
    console.log(error);
  }
})();
